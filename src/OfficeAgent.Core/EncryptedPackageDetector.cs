using System.Text;

namespace OfficeAgent.Core;

/// <summary>
/// Recognises an encrypted Office file: an OLE compound file holding an
/// <c>EncryptedPackage</c> stream.
/// </summary>
/// <remarks>
/// A file protected by a sensitivity label or a password is not an Office package on the
/// outside, so without this check it fails as an unreadable archive and nobody learns why.
/// The file is hostile input like any other: every offset is range-checked and every sector
/// chain is bounded, and anything malformed is simply not an encrypted file, which leaves it
/// to the ordinary package refusal.
/// </remarks>
internal static class EncryptedPackageDetector
{
    private const int HeaderSize = 512;
    private const int DirectoryEntrySize = 128;
    private const int HeaderDifatEntries = 109;
    private const byte StreamEntry = 2;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;

    private static readonly byte[] Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
    private static readonly byte[] StreamName = Encoding.Unicode.GetBytes("EncryptedPackage");

    /// <summary>Whether the bytes are an encrypted Office file.</summary>
    public static bool IsEncrypted(byte[] file, int length)
    {
        if (length < HeaderSize || !HasSignature(file)) return false;

        var shift = ReadUInt16(file, 30);
        if (shift != 9 && shift != 12) return false;

        var sectorSize = 1 << shift;
        var sectors = (length - sectorSize) / sectorSize;
        if (sectors <= 0) return false;

        var fat = FatSectors(file, length, sectorSize, sectors);
        var sector = ReadUInt32(file, 48);

        // A chain cannot be longer than the file has sectors, so a cycle ends here too.
        for (var visited = 0; visited < sectors; visited++)
        {
            if (sector == EndOfChain || sector == FreeSector) return false;
            if (!TryOffset(sector, sectorSize, length, out var offset)) return false;

            for (var entry = 0; entry + DirectoryEntrySize <= sectorSize; entry += DirectoryEntrySize)
                if (IsEncryptedPackageEntry(file, offset + entry))
                    return true;

            if (!TryNext(file, length, fat, sectorSize, sector, out sector)) return false;
        }

        return false;
    }

    private static bool HasSignature(byte[] file)
    {
        for (var i = 0; i < Signature.Length; i++)
            if (file[i] != Signature[i])
                return false;
        return true;
    }

    private static bool IsEncryptedPackageEntry(byte[] file, int entry)
    {
        // The stored length counts the terminating null character.
        if (ReadUInt16(file, entry + 64) != StreamName.Length + 2) return false;
        if (file[entry + 66] != StreamEntry) return false;

        for (var i = 0; i < StreamName.Length; i++)
            if (file[entry + i] != StreamName[i])
                return false;
        return true;
    }

    /// <summary>
    /// The sectors holding the allocation table: up to 109 listed in the header, the rest in a
    /// chain of further sectors, which a file above a few megabytes needs.
    /// </summary>
    private static List<uint> FatSectors(byte[] file, int length, int sectorSize, int sectors)
    {
        var fat = new List<uint>();
        for (var i = 0; i < HeaderDifatEntries; i++)
            Add(fat, ReadUInt32(file, 76 + (i * 4)));

        var entriesPerSector = (sectorSize / 4) - 1;
        var next = ReadUInt32(file, 68);
        for (var visited = 0; visited < sectors && next != EndOfChain && next != FreeSector; visited++)
        {
            if (!TryOffset(next, sectorSize, length, out var offset)) break;
            for (var i = 0; i < entriesPerSector; i++)
                Add(fat, ReadUInt32(file, offset + (i * 4)));
            next = ReadUInt32(file, offset + (entriesPerSector * 4));
        }

        return fat;
    }

    private static void Add(List<uint> fat, uint sector)
    {
        if (sector != FreeSector && sector != EndOfChain) fat.Add(sector);
    }

    private static bool TryNext(byte[] file, int length, List<uint> fat, int sectorSize, uint sector, out uint next)
    {
        next = EndOfChain;
        var perSector = (uint)(sectorSize / 4);
        var index = sector / perSector;
        if (index >= fat.Count) return false;
        if (!TryOffset(fat[(int)index], sectorSize, length, out var offset)) return false;

        next = ReadUInt32(file, offset + (int)(sector % perSector * 4));
        return true;
    }

    /// <summary>The file offset of a whole sector; the header occupies the first sector's place.</summary>
    private static bool TryOffset(uint sector, int sectorSize, int length, out int offset)
    {
        offset = 0;
        var position = ((long)sector + 1) * sectorSize;
        if (position + sectorSize > length) return false;
        offset = (int)position;
        return true;
    }

    private static ushort ReadUInt16(byte[] file, int offset) =>
        (ushort)(file[offset] | (file[offset + 1] << 8));

    private static uint ReadUInt32(byte[] file, int offset) =>
        (uint)(file[offset] | (file[offset + 1] << 8) | (file[offset + 2] << 16) | (file[offset + 3] << 24));
}
