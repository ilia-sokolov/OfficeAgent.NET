using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.SharePoint;

/// <summary>
/// Identifies a Graph drive item by its drive and item ids. A SharePoint
/// connection is not pinned to a single drive, so each registration carries the
/// drive the item lives in alongside the item id.
/// </summary>
/// <param name="DriveId">The Graph drive id the item belongs to.</param>
/// <param name="ItemId">The Graph drive-item id within that drive.</param>
public readonly record struct SharePointItemRef(string DriveId, string ItemId);

/// <summary>
/// Persists the opaque registration id → Graph (drive id, drive-item id) mapping
/// for one SharePoint connection. The store holds only identifiers - never bytes,
/// URLs with embedded tokens, or credentials.
/// </summary>
public interface ISharePointRegistrationStore
{
    /// <summary>Mints a new opaque registration id for a drive item and persists the mapping.</summary>
    Task<string> AddAsync(SharePointItemRef item, CancellationToken cancellationToken = default);

    /// <summary>Returns the drive item for a registration, or <see langword="null"/> when unregistered.</summary>
    Task<SharePointItemRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default);

    /// <summary>Removes a registration. Returns <see langword="false"/> when it did not exist.</summary>
    Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default);
}

/// <summary>
/// Keeps registrations in process memory. The default for tests and short-lived
/// hosts; registrations do not survive a restart. Use
/// <see cref="JsonFileRegistrationStore"/> (or your own store) for durable hosts.
/// </summary>
public sealed class InMemoryRegistrationStore : ISharePointRegistrationStore
{
    private readonly Dictionary<string, SharePointItemRef> _items = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    /// <inheritdoc />
    public Task<string> AddAsync(SharePointItemRef item, CancellationToken cancellationToken = default)
    {
        ValidateItem(item);
        var id = Guid.NewGuid().ToString("N");
        lock (_lock) _items[id] = item;
        return Task.FromResult(id);
    }

    /// <inheritdoc />
    public Task<SharePointItemRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock)
            return Task.FromResult(_items.TryGetValue(registrationId, out var value) ? value : (SharePointItemRef?)null);
    }

    /// <inheritdoc />
    public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        lock (_lock) return Task.FromResult(_items.Remove(registrationId));
    }

    internal static void ValidateItem(SharePointItemRef item)
    {
        if (string.IsNullOrWhiteSpace(item.DriveId))
            throw new ArgumentException("A drive id is required.", nameof(item));
        if (string.IsNullOrWhiteSpace(item.ItemId))
            throw new ArgumentException("A drive item id is required.", nameof(item));
    }
}

/// <summary>
/// Persists registrations in a JSON file with atomic writes, mirroring the
/// filesystem provider's <c>.officeagent/index.json</c> convention. Suitable for
/// single-instance hosts; multi-instance deployments should implement
/// <see cref="ISharePointRegistrationStore"/> over shared storage.
/// </summary>
public sealed class JsonFileRegistrationStore : ISharePointRegistrationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly string _path;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly Dictionary<string, SharePointItemRef> _items;

    /// <summary>Initializes the store over a JSON index file, creating its directory when missing.</summary>
    public JsonFileRegistrationStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An index file path is required.", nameof(path));
        _path = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        _items = Load();
    }

    /// <inheritdoc />
    public async Task<string> AddAsync(SharePointItemRef item, CancellationToken cancellationToken = default)
    {
        InMemoryRegistrationStore.ValidateItem(item);
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var id = Guid.NewGuid().ToString("N");
            _items[id] = item;
            Persist();
            return id;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <inheritdoc />
    public async Task<SharePointItemRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return _items.TryGetValue(registrationId, out var value) ? value : (SharePointItemRef?)null;
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

    private Dictionary<string, SharePointItemRef> Load()
    {
        if (!File.Exists(_path))
            return new Dictionary<string, SharePointItemRef>(StringComparer.Ordinal);
        try
        {
            var model = JsonSerializer.Deserialize<IndexFile>(File.ReadAllText(_path), JsonOptions);
            var loaded = new Dictionary<string, SharePointItemRef>(StringComparer.Ordinal);
            if (model?.Items is { Count: > 0 })
                foreach (var pair in model.Items)
                    loaded[pair.Key] = new SharePointItemRef(pair.Value.DriveId, pair.Value.ItemId);
            return loaded;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            throw new DocumentProviderException(
                ProviderErrorCode.ConfigurationError,
                $"Could not load the SharePoint registration index at '{_path}': {ex.Message}",
                SharePointDocumentProvider.ProviderName, connectionId: "", itemId: null, ex);
        }
    }

    private void Persist()
    {
        var model = new IndexFile
        {
            Version = 2,
            Items = _items.ToDictionary(
                pair => pair.Key,
                pair => new ItemRefRecord { DriveId = pair.Value.DriveId, ItemId = pair.Value.ItemId },
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
        public Dictionary<string, ItemRefRecord> Items { get; set; } = new(StringComparer.Ordinal);
    }

    private sealed class ItemRefRecord
    {
        public string DriveId { get; set; } = string.Empty;
        public string ItemId { get; set; } = string.Empty;
    }
}
