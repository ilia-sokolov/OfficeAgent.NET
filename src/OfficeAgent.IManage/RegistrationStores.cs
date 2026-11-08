using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.IManage;

/// <summary>
/// Identifies an iManage Work document: its library (database), its document number, and a
/// version known to exist when it was registered. A registration addresses the document, so
/// opens and saves always resolve its latest version; <paramref name="Version"/> is only the
/// handle used to look the document up.
/// </summary>
/// <param name="Library">The iManage library (database) id, for example <c>ACTIVE</c>.</param>
/// <param name="DocumentNumber">The document number within the library.</param>
/// <param name="Version">A version of the document that existed when it was registered.</param>
public readonly record struct IManageDocumentRef(string Library, long DocumentNumber, int Version)
{
    /// <summary>Gets the iManage document id of <see cref="Version"/>, <c>LIBRARY!number.version</c>.</summary>
    public string DocumentId => $"{Library}!{DocumentNumber}.{Version}";
}

/// <summary>
/// Persists the opaque registration id → iManage document mapping for one iManage connection.
/// The store holds only identifiers - never bytes, URLs, or credentials.
/// </summary>
public interface IIManageRegistrationStore
{
    /// <summary>Mints a new opaque registration id for a document and persists the mapping.</summary>
    Task<string> AddAsync(IManageDocumentRef document, CancellationToken cancellationToken = default);

    /// <summary>Returns the document for a registration, or <see langword="null"/> when unregistered.</summary>
    Task<IManageDocumentRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default);

    /// <summary>Removes a registration. Returns <see langword="false"/> when it did not exist.</summary>
    Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default);
}

/// <summary>
/// Keeps registrations in process memory. The default for tests and short-lived hosts;
/// registrations do not survive a restart. Use <see cref="JsonFileIManageRegistrationStore"/>
/// (or your own store) for durable hosts.
/// </summary>
public sealed class InMemoryIManageRegistrationStore : IIManageRegistrationStore
{
    private readonly Dictionary<string, IManageDocumentRef> _items = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    /// <inheritdoc />
    public Task<string> AddAsync(IManageDocumentRef document, CancellationToken cancellationToken = default)
    {
        Validate(document);
        var id = Guid.NewGuid().ToString("N");
        lock (_lock) _items[id] = document;
        return Task.FromResult(id);
    }

    /// <inheritdoc />
    public Task<IManageDocumentRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock)
            return Task.FromResult(_items.TryGetValue(registrationId, out var value) ? value : (IManageDocumentRef?)null);
    }

    /// <inheritdoc />
    public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock) return Task.FromResult(_items.Remove(registrationId));
    }

    internal static void Validate(IManageDocumentRef document)
    {
        if (!IManageIds.IsLibrary(document.Library))
            throw new ArgumentException("A valid iManage library id is required.", nameof(document));
        if (document.DocumentNumber <= 0)
            throw new ArgumentException("A positive document number is required.", nameof(document));
        if (document.Version <= 0)
            throw new ArgumentException("A positive document version is required.", nameof(document));
    }
}

/// <summary>
/// Persists registrations in a JSON file with atomic writes, mirroring the filesystem
/// provider's <c>.officeagent/index.json</c> convention. Suitable for single-instance hosts;
/// multi-instance deployments should implement <see cref="IIManageRegistrationStore"/> over
/// shared storage.
/// </summary>
public sealed class JsonFileIManageRegistrationStore : IIManageRegistrationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly string _path;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly Dictionary<string, IManageDocumentRef> _items;

    /// <summary>Initializes the store over a JSON index file, creating its directory when missing.</summary>
    public JsonFileIManageRegistrationStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An index file path is required.", nameof(path));
        _path = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        _items = Load();
    }

    /// <inheritdoc />
    public async Task<string> AddAsync(IManageDocumentRef document, CancellationToken cancellationToken = default)
    {
        InMemoryIManageRegistrationStore.Validate(document);
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var id = Guid.NewGuid().ToString("N");
            _items[id] = document;
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
    public async Task<IManageDocumentRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return _items.TryGetValue(registrationId, out var value) ? value : (IManageDocumentRef?)null;
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
            if (!_items.TryGetValue(registrationId, out var removed)) return false;
            _items.Remove(registrationId);
            try
            {
                Persist();
            }
            catch
            {
                _items[registrationId] = removed;
                throw;
            }
            return true;
        }
        finally
        {
            _lock.Release();
        }
    }

    private Dictionary<string, IManageDocumentRef> Load()
    {
        if (!File.Exists(_path))
            return new Dictionary<string, IManageDocumentRef>(StringComparer.Ordinal);
        try
        {
            var model = JsonSerializer.Deserialize<IndexFile>(File.ReadAllText(_path), JsonOptions);
            var loaded = new Dictionary<string, IManageDocumentRef>(StringComparer.Ordinal);
            if (model?.Items is { Count: > 0 })
                foreach (var pair in model.Items)
                    loaded[pair.Key] = new IManageDocumentRef(pair.Value.Library, pair.Value.DocumentNumber, pair.Value.Version);
            return loaded;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Could not load the iManage registration index at '{_path}': {ex.Message}",
                IManageDocumentProvider.ProviderName, connectionId: "", itemId: null, ex);
        }
    }

    private void Persist()
    {
        var model = new IndexFile
        {
            Version = 1,
            Items = _items.ToDictionary(
                pair => pair.Key,
                pair => new DocumentRecord
                {
                    Library = pair.Value.Library,
                    DocumentNumber = pair.Value.DocumentNumber,
                    Version = pair.Value.Version
                },
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
        public Dictionary<string, DocumentRecord> Items { get; set; } = new(StringComparer.Ordinal);
    }

    private sealed class DocumentRecord
    {
        public string Library { get; set; } = string.Empty;
        public long DocumentNumber { get; set; }
        public int Version { get; set; }
    }
}
