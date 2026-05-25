namespace OneDriveCopier;

// ─────────────────────────────────────────────────────────────────────────────
// Strongly-typed settings — bound from appsettings.json via
// Microsoft.Extensions.Configuration.Binder.
// ─────────────────────────────────────────────────────────────────────────────

internal sealed class AppSettings
{
    public AzureAdSettings     AzureAd     { get; set; } = new();

    /// <summary>
    /// "OneDrive" | "SharePoint" | "AzureBlob"
    /// Toggle with one value — all other settings are in their own blocks.
    /// </summary>
    public string              SourceMode  { get; set; } = "OneDrive";
    public OneDriveSettings    OneDrive    { get; set; } = new();
    public SharePointSettings  SharePoint  { get; set; } = new();
    public AzureBlobSettings   AzureBlob   { get; set; } = new();
    public DestinationSettings Destination { get; set; } = new();
    public CopyOptionsSettings CopyOptions { get; set; } = new();
    public RetryPolicySettings RetryPolicy { get; set; } = new();
    public LoggingSettings     Logging     { get; set; } = new();
}

// ── Azure AD / auth ───────────────────────────────────────────────────────────

internal sealed class AzureAdSettings
{
    public string TenantId     { get; set; } = string.Empty;
    public string ClientId     { get; set; } = string.Empty;

    /// <summary>
    /// "ClientCredentials" — app-only, no browser. Default. Use for Autosys.
    ///     Requires ClientSecret.
    ///     App needs APPLICATION permissions (admin-consented):
    ///       Files.Read.All, Sites.Read.All
    ///
    /// "Interactive" — browser pop-up on first run; MSAL token cached to disk
    ///     for subsequent runs until the token expires.
    ///     App needs DELEGATED permissions: Files.Read, Files.Read.All, Sites.Read.All
    ///     App registration: Mobile/desktop platform, redirect http://localhost,
    ///     Allow public client flows = Yes.
    ///
    /// Can be overridden at runtime with --auth-mode CLI flag.
    /// </summary>
    public string AuthMode     { get; set; } = "ClientCredentials";

    /// <summary>Required when AuthMode = "ClientCredentials".</summary>
    public string ClientSecret { get; set; } = string.Empty;

    public InteractiveLoginSettings InteractiveLogin { get; set; } = new();
}

internal sealed class InteractiveLoginSettings
{
    /// <summary>
    /// Directory for the MSAL token cache file.
    /// Leave blank to place the cache next to the exe.
    /// </summary>
    public string TokenCacheDir { get; set; } = string.Empty;

    /// <summary>Must match the redirect URI registered in Azure AD.</summary>
    public string RedirectUri   { get; set; } = "http://localhost";
}

// ── Source settings ───────────────────────────────────────────────────────────

internal sealed class OneDriveSettings
{
    /// <summary>
    /// UPN or object ID of the drive owner.
    /// Required when AuthMode = "ClientCredentials".
    /// Optional when AuthMode = "Interactive" — leave blank to use /me/drive.
    /// </summary>
    public string UserId         { get; set; } = string.Empty;

    /// <summary>Folder path inside drive root, excluding the date folder. e.g. "Reports"</summary>
    public string BaseFolderPath { get; set; } = string.Empty;
}

internal sealed class SharePointSettings
{
    /// <summary>e.g. "contoso.sharepoint.com"</summary>
    public string Hostname       { get; set; } = string.Empty;

    /// <summary>
    /// Relative site path without leading slash. e.g. "sites/FinanceTeam"
    /// Combined with Hostname: GET /sites/{Hostname}:/{SitePath}
    /// </summary>
    public string SitePath       { get; set; } = string.Empty;

    /// <summary>Document library display name. e.g. "Documents"</summary>
    public string LibraryName    { get; set; } = string.Empty;

    /// <summary>Folder path inside library root, excluding the date folder. e.g. "Reports"</summary>
    public string BaseFolderPath { get; set; } = string.Empty;
}

internal sealed class AzureBlobSettings
{
    /// <summary>
    /// "ConnectionString" — full Azure Storage connection string (simplest).
    /// "ClientCredentials" — AAD app identity; reuses AzureAd block credentials.
    ///     App needs Storage Blob Data Reader role on the container.
    /// "ManagedIdentity" — VM/App Service managed identity; no secrets.
    /// </summary>
    public string AuthMode         { get; set; } = "ConnectionString";
    public string ConnectionString { get; set; } = string.Empty;
    public string AccountUri       { get; set; } = string.Empty;
    public string ContainerName    { get; set; } = string.Empty;

    /// <summary>
    /// Virtual folder prefix excluding date folder.
    /// e.g. "finance/reports" → blobs at "finance/reports/20260522/"
    /// Leave blank for "20260522/" at container root.
    /// </summary>
    public string BaseFolderPrefix { get; set; } = string.Empty;
}

// ── Operation settings ────────────────────────────────────────────────────────

internal sealed class DestinationSettings
{
    public string RootPath { get; set; } = string.Empty;
}

internal sealed class CopyOptionsSettings
{
    public bool   Overwrite         { get; set; } = false;
    public string FileFilter        { get; set; } = "*.xls,*.xlsx";
    public int    MaxRetries        { get; set; } = 3;
    public int    RetryDelaySeconds { get; set; } = 1;
}

internal sealed class RetryPolicySettings
{
    /// <summary>
    /// Maximum number of full-run attempts before giving up.
    /// Default 30 attempts × 10-minute delay = up to 5 hours total.
    /// Used by Autosys-style outer retry loop distinct from per-file copy retries.
    /// </summary>
    public int MaxAttempts       { get; set; } = 30;

    /// <summary>Minutes to wait between attempts.</summary>
    public int DelayMinutes      { get; set; } = 10;

    /// <summary>Hard ceiling in hours — stops retrying even if MaxAttempts not reached.</summary>
    public double MaxDurationHours { get; set; } = 5.0;
}

internal sealed class LoggingSettings
{
    /// <summary>Leave blank to write logs next to the exe.</summary>
    public string LogDirectory { get; set; } = string.Empty;
}
