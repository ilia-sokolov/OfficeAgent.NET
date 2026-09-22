using System.Security.Cryptography;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// Runs one provider save or create and classifies how it failed, so every workflow reports a
/// storage outcome the same way.
/// </summary>
/// <remarks>
/// Only three outcomes are honest after a write call fails. The provider may prove nothing was
/// written (a pre-write refusal or <see cref="ProviderErrorCode.WriteRejected"/>); it may prove the
/// bytes were stored but not registered (<see cref="ProviderErrorCode.RegistrationFailed"/>); or
/// the engine cannot know. Everything a provider did not classify as certain is the third case:
/// an <see cref="ProviderErrorCode.IO"/> failure, a transport error, a timeout, and a cancellation
/// that arrives once the call has begun. Reporting any of them as "nothing written" would invite
/// a blind retry that duplicates or overwrites a write that did land.
/// </remarks>
internal static class StorageWrite
{
    /// <summary>Codes a provider raises only before anything was stored.</summary>
    internal static bool ProvesNothingWritten(ProviderErrorCode code) => code
        is ProviderErrorCode.NotFound
        or ProviderErrorCode.AccessDenied
        or ProviderErrorCode.ContentTooLarge
        or ProviderErrorCode.ExtensionNotAllowed
        or ProviderErrorCode.VersionConflict
        or ProviderErrorCode.InvalidArgument
        or ProviderErrorCode.ConfigurationError
        or ProviderErrorCode.AlreadyExists
        or ProviderErrorCode.WriteRejected;

    /// <summary>The stable code a provider failure is reported with, on every surface.</summary>
    internal static string WireCode(ProviderErrorCode code) => code switch
    {
        ProviderErrorCode.NotFound => ToolErrorCodes.NotFound,
        ProviderErrorCode.AccessDenied => ToolErrorCodes.AccessDenied,
        ProviderErrorCode.ContentTooLarge => ToolErrorCodes.ContentTooLarge,
        ProviderErrorCode.ExtensionNotAllowed => ToolErrorCodes.ExtensionNotAllowed,
        ProviderErrorCode.VersionConflict => ToolErrorCodes.VersionConflict,
        ProviderErrorCode.InvalidArgument => ToolErrorCodes.InvalidArgument,
        ProviderErrorCode.ConfigurationError => ToolErrorCodes.ConfigurationError,
        ProviderErrorCode.IO => ToolErrorCodes.IOError,
        ProviderErrorCode.AlreadyExists => ToolErrorCodes.AlreadyExists,
        ProviderErrorCode.WriteRejected => ToolErrorCodes.WriteRejected,
        ProviderErrorCode.OutcomeUnknown => ToolErrorCodes.OutcomeUnknown,
        ProviderErrorCode.RegistrationFailed => ToolErrorCodes.RegistrationFailed,
        _ => ToolErrorCodes.ProviderError
    };

    public static async Task<DocumentReference> RunAsync(
        IDocumentProvider provider,
        string? itemId,
        string? outputName,
        byte[] content,
        Func<CancellationToken, Task<DocumentReference>> write,
        CancellationToken cancellationToken)
    {
        // Before the call nothing can have been written, so a cancellation here is plain.
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            return await write(cancellationToken).ConfigureAwait(false);
        }
        catch (DocumentProviderException ex) when (
            ex.Code is ProviderErrorCode.RegistrationFailed &&
            ex is not DocumentWriteRecoveryException &&
            !string.IsNullOrEmpty(outputName))
        {
            // A provider that reported the registration failure without the locator: the engine
            // knows the name it asked for and the exact bytes, so the caller can still find and
            // register the stored output instead of writing it a second time.
            throw new DocumentRegistrationFailedException(
                ex.Message, ex.Provider, ex.ConnectionId, ex.ItemId, outputName!, Sha256(content), ex);
        }
        catch (DocumentProviderException ex) when (
            ProvesNothingWritten(ex.Code) ||
            ex.Code is ProviderErrorCode.RegistrationFailed ||
            ex is DocumentWriteRecoveryException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new DocumentWriteOutcomeUnknownException(
                "The storage may already hold this output: the write failed, timed out or was cancelled " +
                "after it began. Do not retry blindly and do not report the document as unchanged; reopen " +
                "the destination and compare it with the intended output first.",
                provider.Provider, provider.ConnectionId, itemId, outputName, Sha256(content), ex);
        }
    }

    private static string Sha256(byte[] content)
    {
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(content).Select(b => b.ToString("x2")));
    }
}
