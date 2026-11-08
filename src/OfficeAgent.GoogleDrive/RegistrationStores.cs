using System.Text.Json;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.GoogleDrive;

/// <summary>
/// Identifies a Google Drive file. Drive file ids are unique across My Drive and shared
/// drives, so the id alone addresses the file.
/// </summary>
/// <param name="FileId">The Drive file id.</param>
public readonly record struct GoogleDriveFileRef(string FileId);

/// <summary>
/// Persists the opaque registration id → Drive file id mapping for one Google Drive
/// connection. The store holds only identifiers - never bytes, URLs, or credentials.
/// </summary>
public interface IGoogleDriveRegistrationStore
{
    /// <summary>Mints a new opaque registration id for a Drive file and persists the mapping.</summary>
    Task<string> AddAsync(GoogleDriveFileRef file, CancellationToken cancellationToken = default);

    /// <summary>Returns the Drive file for a registration, or <see langword="null"/> when unregistered.</summary>
    Task<GoogleDriveFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default);

    /// <summary>Removes a registration. Returns <see langword="false"/> when it did not exist.</summary>
    Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default);
}

/// <summary>
/// Keeps Google Drive registrations in process memory. The default for tests and
/// short-lived hosts; registrations do not survive a restart. Use
/// <see cref="JsonFileGoogleDriveRegistrationStore"/> (or your own store) for durable hosts.
/// </summary>
public sealed class InMemoryGoogleDriveRegistrationStore : IGoogleDriveRegistrationStore
{
    private readonly Dictionary<string, GoogleDriveFileRef> _items = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    /// <inheritdoc />
    public Task<string> AddAsync(GoogleDriveFileRef file, CancellationToken cancellationToken = default)
    {
        Validate(file);
        var id = Guid.NewGuid().ToString("N");
        lock (_lock) _items[id] = file;
        return Task.FromResult(id);
    }

    /// <inheritdoc />
    public Task<GoogleDriveFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock)
            return Task.FromResult(_items.TryGetValue(registrationId, out var value) ? value : (GoogleDriveFileRef?)null);
    }

    /// <inheritdoc />
    public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock) return Task.FromResult(_items.Remove(registrationId));
    }

    internal static void Validate(GoogleDriveFileRef file)
    {
        if (string.IsNullOrWhiteSpace(file.FileId))
            throw new ArgumentException("A Drive file id is required.", nameof(file));
    }
}

/// <summary>
/// Persists Google Drive registrations in a JSON file with atomic writes. Suitable for
/// single-instance hosts; multi-instance deployments should implement
/// <see cref="IGoogleDriveRegistrationStore"/> over shared storage.
/// </summary>
public sealed class JsonFileGoogleDriveRegistrationStore : IGoogleDriveRegistrationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _path;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly Dictionary<string, GoogleDriveFileRef> _items;

    /// <summary>Initializes the store over a JSON index file, creating its directory when missing.</summary>
    public JsonFileGoogleDriveRegistrationStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An index file path is required.", nameof(path));
        _path = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        _items = Load();
    }

    /// <inheritdoc />
    public async Task<string> AddAsync(GoogleDriveFileRef file, CancellationToken cancellationToken = default)
    {
        InMemoryGoogleDriveRegistrationStore.Validate(file);
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var id = Guid.NewGuid().ToString("N");
            _items[id] = file;
            try
            {
                Persist();
            }
            catch
            {
                _items.Remove(id);
                throw;
            }
            return id;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <inheritdoc />
    public async Task<GoogleDriveFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return _items.TryGetValue(registrationId, out var value) ? value : (GoogleDriveFileRef?)null;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <inheritdoc />
    public async Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!_items.Remove(registrationId)) return false;
            Persist();
            return true;
        }
        finally
        {
            _lock.Release();
        }
    }

    private Dictionary<string, GoogleDriveFileRef> Load()
    {
        if (!File.Exists(_path))
            return new Dictionary<string, GoogleDriveFileRef>(StringComparer.Ordinal);
        try
        {
            var model = JsonSerializer.Deserialize<IndexFile>(File.ReadAllText(_path), JsonOptions);
            var loaded = new Dictionary<string, GoogleDriveFileRef>(StringComparer.Ordinal);
            if (model?.Items is { Count: > 0 })
                foreach (var pair in model.Items)
                    loaded[pair.Key] = new GoogleDriveFileRef(pair.Value.FileId);
            return loaded;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Could not load the Google Drive registration index at '{_path}': {ex.Message}",
                GoogleDriveDocumentProvider.ProviderName, connectionId: "", itemId: null, ex);
        }
    }

    private void Persist()
    {
        var model = new IndexFile
        {
            Version = 1,
            Items = _items.ToDictionary(
                pair => pair.Key,
                pair => new FileRecord { FileId = pair.Value.FileId },
                StringComparer.Ordinal)
        };
        var temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(model, JsonOptions));
            if (File.Exists(_path))
                File.Replace(temporary, _path, null);
            else
                File.Move(temporary, _path);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    private sealed class IndexFile
    {
        public int Version { get; set; }
        public Dictionary<string, FileRecord> Items { get; set; } = new(StringComparer.Ordinal);
    }

    private sealed class FileRecord
    {
        public string FileId { get; set; } = string.Empty;
    }
}
