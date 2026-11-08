using System.Text.Json;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Box;

/// <summary>
/// Identifies a Box file. Box file ids are unique within the Box service, so the id alone
/// addresses the file.
/// </summary>
/// <param name="FileId">The Box file id.</param>
public readonly record struct BoxFileRef(string FileId);

/// <summary>
/// Persists the opaque registration id → Box file id mapping for one Box
/// connection. The store holds only identifiers - never bytes, URLs, or credentials.
/// </summary>
public interface IBoxRegistrationStore
{
    /// <summary>Mints a new opaque registration id for a Box file and persists the mapping.</summary>
    Task<string> AddAsync(BoxFileRef file, CancellationToken cancellationToken = default);

    /// <summary>Returns the Box file for a registration, or <see langword="null"/> when unregistered.</summary>
    Task<BoxFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default);

    /// <summary>Removes a registration. Returns <see langword="false"/> when it did not exist.</summary>
    Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default);
}

/// <summary>
/// Keeps Box registrations in process memory. The default for tests and
/// short-lived hosts; registrations do not survive a restart. Use
/// <see cref="JsonFileBoxRegistrationStore"/> (or your own store) for durable hosts.
/// </summary>
public sealed class InMemoryBoxRegistrationStore : IBoxRegistrationStore
{
    private readonly Dictionary<string, BoxFileRef> _items = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    /// <inheritdoc />
    public Task<string> AddAsync(BoxFileRef file, CancellationToken cancellationToken = default)
    {
        Validate(file);
        var id = Guid.NewGuid().ToString("N");
        lock (_lock) _items[id] = file;
        return Task.FromResult(id);
    }

    /// <inheritdoc />
    public Task<BoxFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock)
            return Task.FromResult(_items.TryGetValue(registrationId, out var value) ? value : (BoxFileRef?)null);
    }

    /// <inheritdoc />
    public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock) return Task.FromResult(_items.Remove(registrationId));
    }

    internal static void Validate(BoxFileRef file)
    {
        if (string.IsNullOrWhiteSpace(file.FileId))
            throw new ArgumentException("A Box file id is required.", nameof(file));
    }
}

/// <summary>
/// Persists Box registrations in a JSON file with atomic writes. Suitable for
/// single-instance hosts; multi-instance deployments should implement
/// <see cref="IBoxRegistrationStore"/> over shared storage.
/// </summary>
public sealed class JsonFileBoxRegistrationStore : IBoxRegistrationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _path;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly Dictionary<string, BoxFileRef> _items;

    /// <summary>Initializes the store over a JSON index file, creating its directory when missing.</summary>
    public JsonFileBoxRegistrationStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An index file path is required.", nameof(path));
        _path = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        _items = Load();
    }

    /// <inheritdoc />
    public async Task<string> AddAsync(BoxFileRef file, CancellationToken cancellationToken = default)
    {
        InMemoryBoxRegistrationStore.Validate(file);
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
    public async Task<BoxFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return _items.TryGetValue(registrationId, out var value) ? value : (BoxFileRef?)null;
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

    private Dictionary<string, BoxFileRef> Load()
    {
        if (!File.Exists(_path))
            return new Dictionary<string, BoxFileRef>(StringComparer.Ordinal);
        try
        {
            var model = JsonSerializer.Deserialize<IndexFile>(File.ReadAllText(_path), JsonOptions);
            var loaded = new Dictionary<string, BoxFileRef>(StringComparer.Ordinal);
            if (model?.Items is { Count: > 0 })
                foreach (var pair in model.Items)
                    loaded[pair.Key] = new BoxFileRef(pair.Value.FileId);
            return loaded;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Could not load the Box registration index at '{_path}': {ex.Message}",
                BoxDocumentProvider.ProviderName, connectionId: "", itemId: null, ex);
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
