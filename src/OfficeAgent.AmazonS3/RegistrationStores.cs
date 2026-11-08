namespace OfficeAgent.ObjectStorage;

// The shared registration indexes (src/Shared/ObjectStorage) serve as this package's stores.
internal sealed partial class InMemoryRegistrationIndex : OfficeAgent.AmazonS3.IAmazonS3RegistrationStore
{
}

internal sealed partial class JsonFileRegistrationIndex : OfficeAgent.AmazonS3.IAmazonS3RegistrationStore
{
}
