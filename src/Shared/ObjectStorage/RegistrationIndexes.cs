// Compiled into OfficeAgent.AzureBlob and OfficeAgent.AmazonS3 as linked source. Each package
// declares the other half of these partial classes, adding its own public store interface.

using System.Text.Json;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.ObjectStorage;

/// <summary>
/// Keeps registrations (opaque id → full object key) in process memory. The default; it does not
/// survive a restart.
/// </summary>
internal sealed partial class InMemoryRegistrationIndex
{
    private readonly Dictionary<string, string> _items = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    public Task<string> AddAsync(string key, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(key)) throw new ArgumentException("A key is required.", nameof(key));
        var id = Guid.NewGuid().ToString("N");
        lock (_lock) _items[id] = key;
        return Task.FromResult(id);
    }

    public Task<string?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock)
            return Task.FromResult(_items.TryGetValue(registrationId, out var key) ? key : null);
    }

    public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock) return Task.FromResult(_items.Remove(registrationId));
    }
}

/// <summary>
/// Persists registrations in a JSON file with atomic replacement, like the filesystem provider's
/// <c>.officeagent/index.json</c>. For a single host process; hosts that run several instances
/// implement the package's registration-store interface over shared storage instead.
/// </summary>
internal sealed partial class JsonFileRegistrationIndex
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _path;
    private readonly string _provider;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly Dictionary<string, string> _items;

    public JsonFileRegistrationIndex(string path, string provider)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An index file path is required.", nameof(path));
        _provider = provider;
        _path = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        _items = Load();
    }

    public async Task<string> AddAsync(string key, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(key)) throw new ArgumentException("A key is required.", nameof(key));
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var id = Guid.NewGuid().ToString("N");
            _items[id] = key;
            try
            {
                Persist();
            }
            catch
            {
                // Memory and disk must agree: an id that was not persisted must not resolve.
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

    public async Task<string?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return _items.TryGetValue(registrationId, out var key) ? key : null;
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!_items.TryGetValue(registrationId, out var key)) return false;
            _items.Remove(registrationId);
            try
            {
                Persist();
            }
            catch
            {
                _items[registrationId] = key;
                throw;
            }
            return true;
        }
        finally
        {
            _lock.Release();
        }
    }

    private Dictionary<string, string> Load()
    {
        if (!File.Exists(_path))
            return new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            var model = JsonSerializer.Deserialize<IndexFile>(File.ReadAllText(_path), JsonOptions);
            return model?.Items is { Count: > 0 }
                ? new Dictionary<string, string>(model.Items, StringComparer.Ordinal)
                : new Dictionary<string, string>(StringComparer.Ordinal);
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Could not load the registration index at '{_path}': {ex.Message}",
                _provider, connectionId: "", itemId: null, ex);
        }
    }

    private void Persist()
    {
        var model = new IndexFile { Version = 1, Items = new Dictionary<string, string>(_items, StringComparer.Ordinal) };
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
        public Dictionary<string, string> Items { get; set; } = new(StringComparer.Ordinal);
    }
}
