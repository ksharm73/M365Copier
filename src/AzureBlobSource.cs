using Azure.Identity;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;

namespace OneDriveCopier;

/// <summary>
/// IFileSource implementation that reads from an Azure Blob Storage container.
///
/// Blob path convention (mirrors OneDrive/SharePoint folder structure):
///   {BaseFolderPrefix}/{dateFolderName}/{fileName}
///
/// Example with BaseFolderPrefix = "finance/reports", dateFolderName = "20260522":
///   finance/reports/20260522/Q1_Risk.xlsx
///   finance/reports/20260522/Q1_Sales.xlsx
///
/// If BaseFolderPrefix is blank, blobs are listed under:
///   20260522/Q1_Risk.xlsx
///
/// ── Authentication options ────────────────────────────────────────────────
///
///   AzureBlob.AuthMode = "ConnectionString"  (default — simplest)
///     Uses a full Azure Storage connection string.
///     Best for: local dev, manual ad-hoc runs, scenarios without AAD.
///     Config: AzureBlob.ConnectionString
///
///   AzureBlob.AuthMode = "ClientCredentials"
///     Authenticates as an AAD application (same TenantId/ClientId/ClientSecret
///     from the AzureAd config block).  No connection string needed.
///     The app registration needs Storage Blob Data Reader role on the container
///     or storage account.
///     Best for: Autosys / service accounts without a managed identity.
///     Config: AzureAd.TenantId, AzureAd.ClientId, AzureAd.ClientSecret
///             AzureBlob.AccountUri
///
///   AzureBlob.AuthMode = "ManagedIdentity"
///     Uses the VM or App Service managed identity — no secrets stored anywhere.
///     The managed identity needs Storage Blob Data Reader role.
///     Best for: production workloads running inside Azure.
///     Config: AzureBlob.AccountUri
///
/// Note: "Interactive" auth is NOT supported for Azure Blob Storage.
/// The Blob SDK uses Azure.Core credentials, not MSAL browser flows.
/// Use ConnectionString for manual runs instead.
/// </summary>
internal sealed class AzureBlobSource : IFileSource
{
    private readonly BlobContainerClient _container;
    private readonly AzureBlobSettings  _cfg;

    public string SourceLabel =>
        $"Azure Blob (container: {_cfg.ContainerName}, auth: {_cfg.AuthMode})";

    /// <param name="cfg">AzureBlob configuration block from appsettings.json.</param>
    /// <param name="azureAdCfg">
    /// AzureAd block — only used when AuthMode = "ClientCredentials" to reuse the
    /// same app registration credentials, avoiding duplication in config.
    /// </param>
    public AzureBlobSource(AzureBlobSettings cfg, AzureAdSettings azureAdCfg)
    {
        _cfg = cfg;
        ValidateConfig(cfg, azureAdCfg);
        _container = BuildContainerClient(cfg, azureAdCfg);
    }

    // ── IFileSource ──────────────────────────────────────────────────────────

    public async Task<IReadOnlyList<RemoteFile>> ListFilesAsync(
        string dateFolderName,
        string[] filters,
        CancellationToken ct = default)
    {
        string prefix = BuildPrefix(dateFolderName);

        var result  = new List<RemoteFile>();
        bool exists = false;

        // ListBlobsFlatAsync streams pages of blobs matching the prefix.
        // There is no explicit "folder exists" check in Blob Storage —
        // an empty result for a well-formed prefix means the folder is absent.
        await foreach (BlobItem blob in _container
            .GetBlobsAsync(prefix: prefix, cancellationToken: ct))
        {
            exists = true;  // at least one blob found → prefix/folder exists

            // Strip the folder prefix to get just the filename (or relative path)
            string relativePath = blob.Name[prefix.Length..].TrimStart('/');

            if (string.IsNullOrEmpty(relativePath)) continue;  // skip "folder" marker blobs
            if (!MatchesFilter(relativePath, filters)) continue;

            result.Add(new RemoteFile
            {
                Name         = Path.GetFileName(relativePath),
                RelativePath = relativePath,
                SizeBytes    = blob.Properties.ContentLength ?? 0
            });
        }

        if (!exists)
            throw new FolderNotFoundException(
                $"Azure Blob prefix not found or empty: {_cfg.ContainerName}/{prefix}  " +
                $"(Account: {_cfg.AccountUri}, BaseFolderPrefix: \"{_cfg.BaseFolderPrefix}\")");

        return result;
    }

    public async Task<Stream> OpenReadAsync(
        RemoteFile file,
        string dateFolderName,
        CancellationToken ct = default)
    {
        // Full blob path = prefix + relative path recorded during listing
        string blobName = BuildPrefix(dateFolderName) + file.RelativePath;

        BlobClient blobClient = _container.GetBlobClient(blobName);

        // Download to a MemoryStream so the caller gets a fully buffered,
        // seekable stream — identical behaviour to the Graph API sources.
        var ms = new MemoryStream();
        await blobClient.DownloadToAsync(ms, ct);
        ms.Position = 0;
        return ms;
    }

    // ── Private: client construction ─────────────────────────────────────────

    private static BlobContainerClient BuildContainerClient(
        AzureBlobSettings cfg,
        AzureAdSettings   azureAdCfg)
    {
        BlobServiceClient serviceClient = cfg.AuthMode.Trim().ToLowerInvariant() switch
        {
            "connectionstring" => new BlobServiceClient(cfg.ConnectionString),

            "clientcredentials" => new BlobServiceClient(
                new Uri(cfg.AccountUri),
                new ClientSecretCredential(
                    azureAdCfg.TenantId,
                    azureAdCfg.ClientId,
                    azureAdCfg.ClientSecret)),

            "managedidentity" => new BlobServiceClient(
                new Uri(cfg.AccountUri),
                new ManagedIdentityCredential()),

            _ => throw new InvalidOperationException(
                $"Unknown AzureBlob.AuthMode '{cfg.AuthMode}'. " +
                "Valid values: ConnectionString, ClientCredentials, ManagedIdentity")
        };

        return serviceClient.GetBlobContainerClient(cfg.ContainerName);
    }

    // ── Private: helpers ─────────────────────────────────────────────────────

    /// <summary>
    /// Builds the blob prefix for the given date folder.
    /// Always ends with "/" so listing is scoped to files inside the folder.
    ///   BaseFolderPrefix = "finance/reports", date = "20260522"
    ///   → "finance/reports/20260522/"
    ///   BaseFolderPrefix = "", date = "20260522"
    ///   → "20260522/"
    /// </summary>
    private string BuildPrefix(string dateFolderName)
    {
        string base_ = _cfg.BaseFolderPrefix.Trim('/').Trim('\\');
        string prefix = string.IsNullOrEmpty(base_)
            ? $"{dateFolderName}/"
            : $"{base_}/{dateFolderName}/";
        return prefix;
    }

    private static bool MatchesFilter(string filePath, string[] filters)
    {
        if (filters.Length == 1 && filters[0] == "*.*") return true;
        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        return filters.Any(f =>
        {
            string filterExt = f.StartsWith("*.") ? f[1..].ToLowerInvariant() : f.ToLowerInvariant();
            return ext == filterExt || f == "*.*";
        });
    }

    private static void ValidateConfig(AzureBlobSettings cfg, AzureAdSettings azureAdCfg)
    {
        if (string.IsNullOrWhiteSpace(cfg.ContainerName))
            throw new InvalidOperationException(
                "AzureBlob.ContainerName is not configured in appsettings.json.");

        switch (cfg.AuthMode.Trim().ToLowerInvariant())
        {
            case "connectionstring":
                if (string.IsNullOrWhiteSpace(cfg.ConnectionString))
                    throw new InvalidOperationException(
                        "AzureBlob.ConnectionString is required when AzureBlob.AuthMode = \"ConnectionString\".");
                break;

            case "clientcredentials":
                if (string.IsNullOrWhiteSpace(cfg.AccountUri))
                    throw new InvalidOperationException(
                        "AzureBlob.AccountUri is required when AzureBlob.AuthMode = \"ClientCredentials\".");
                if (string.IsNullOrWhiteSpace(azureAdCfg.ClientSecret))
                    throw new InvalidOperationException(
                        "AzureAd.ClientSecret is required when AzureBlob.AuthMode = \"ClientCredentials\".");
                break;

            case "managedidentity":
                if (string.IsNullOrWhiteSpace(cfg.AccountUri))
                    throw new InvalidOperationException(
                        "AzureBlob.AccountUri is required when AzureBlob.AuthMode = \"ManagedIdentity\".");
                break;
        }
    }
}
