namespace OfficeAgent.ObjectStorage;

// The shared registration indexes (src/Shared/ObjectStorage) serve as this package's stores.
internal sealed partial class InMemoryRegistrationIndex : OfficeAgent.AzureBlob.IAzureBlobRegistrationStore
{
}

internal sealed partial class JsonFileRegistrationIndex : OfficeAgent.AzureBlob.IAzureBlobRegistrationStore
{
}
