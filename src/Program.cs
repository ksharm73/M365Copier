using Microsoft.Extensions.Configuration;
using OneDriveCopier;

/// <summary>
/// OneDrive / SharePoint / Azure Blob → Network Drive File Copier  v4.0
///
/// Graph API access via raw HttpClient — no Microsoft.Graph SDK.
///
/// Usage:
///   OneDriveCopier.exe --folder-name 20260522 [--auth-mode Interactive]
///
/// SOURCE MODES  (SourceMode in appsettings.json):
///   OneDrive    — user's OneDrive via Graph REST API
///   SharePoint  — SharePoint document library via Graph REST API
///   AzureBlob   — Azure Blob Storage container (own SDK, no Graph)
///
/// AUTH MODES for OneDrive + SharePoint  (AzureAd.AuthMode):
///   ClientCredentials  — app-only, no browser (default — Autosys/scheduler)
///   Interactive        — browser login, MSAL token cached to disk
///
/// AUTH MODES for AzureBlob  (AzureBlob.AuthMode):
///   ConnectionString / ClientCredentials / ManagedIdentity
///
/// RETRY POLICY  (RetryPolicy in appsettings.json):
///   MaxAttempts      — max full-run attempts           (default: 30)
///   DelayMinutes     — wait between attempts           (default: 10)
///   MaxDurationHours — hard ceiling in hours           (default: 5.0)
/// </summary>
class Program
{
    static async Task<int> Main(string[] args)
    {
        var parser = new ArgumentParser(args);

        //if (parser.HasFlag("--help") || parser.HasFlag("-h") || args.Length == 0)
        //{
        //    PrintUsage();
        //    return 0;
        //}

        // ── 1. Load configuration ───────────────────────────────────────────
        string appDir = AppContext.BaseDirectory;

        IConfiguration configuration = new ConfigurationBuilder()
            .SetBasePath(appDir)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .Build();

        var settings = new AppSettings();
        configuration.Bind(settings);

        // ── 2. Parse CLI arguments ──────────────────────────────────────────
        string? folderName = "20260524";//parser.GetValue("--folder-name");
        string? destOverride     = parser.GetValue("--dest");
        string? logDirOverride   = parser.GetValue("--log-dir");
        string? authModeOverride = parser.GetValue("--auth-mode");

        if (string.IsNullOrWhiteSpace(folderName))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine("ERROR: --folder-name is required (e.g. --folder-name 20260522)");
            Console.ResetColor();
            PrintUsage();
            return ExitCodes.InvalidArguments;
        }

        string destRoot = !string.IsNullOrWhiteSpace(destOverride)
            ? destOverride
            : settings.Destination.RootPath;

        if (string.IsNullOrWhiteSpace(destRoot))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine(
                "ERROR: Destination.RootPath is not set in appsettings.json " +
                "and --dest was not provided.");
            Console.ResetColor();
            return ExitCodes.InvalidArguments;
        }

        string logDir = !string.IsNullOrWhiteSpace(logDirOverride)
            ? logDirOverride
            : !string.IsNullOrWhiteSpace(settings.Logging.LogDirectory)
                ? settings.Logging.LogDirectory
                : appDir;

        string[] filters = string.IsNullOrWhiteSpace(settings.CopyOptions.FileFilter)
            ? ["*.*"]
            : settings.CopyOptions.FileFilter
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        // ── 3. Initialise logger ────────────────────────────────────────────
        using var logger = new FileLogger(logDir);
        logger.Info($"Source mode  : {settings.SourceMode}");

        // ── 4. Resolve auth mode ────────────────────────────────────────────
        string resolvedAuthMode = (authModeOverride ?? settings.AzureAd.AuthMode).Trim();
        bool isInteractive = resolvedAuthMode.Equals(
            "Interactive", StringComparison.OrdinalIgnoreCase);

        // ── 5. Build GraphHttpClient (only for OneDrive / SharePoint) ───────
        string sourceModeLower = settings.SourceMode.Trim().ToLowerInvariant();
        bool needsGraph = sourceModeLower is "onedrive" or "sharepoint";

        GraphHttpClient? graphHttp = null;
        if (needsGraph)
        {
            try
            {
                graphHttp = GraphHttpClient.Create(
                    settings.AzureAd, authModeOverride, logger);
            }
            catch (InvalidOperationException ex)
            {
                logger.Error($"Auth configuration error: {ex.Message}");
                return ExitCodes.InvalidArguments;
            }
        }
        else if (authModeOverride is not null)
        {
            logger.Warn($"--auth-mode is ignored for SourceMode \"{settings.SourceMode}\". " +
                        "Configure AzureBlob.AuthMode in appsettings.json instead.");
        }

        // ── 6. Build file source ────────────────────────────────────────────
        IFileSource source;
        try
        {
            source = sourceModeLower switch
            {
                "onedrive"   => new OneDriveSource(graphHttp!, settings.OneDrive, isInteractive),
                "sharepoint" => new SharePointSource(graphHttp!, settings.SharePoint),
                "azureblob"  => new AzureBlobSource(settings.AzureBlob, settings.AzureAd),
                _ => throw new InvalidOperationException(
                    $"Unknown SourceMode '{settings.SourceMode}'. " +
                    "Valid values: OneDrive, SharePoint, AzureBlob")
            };
        }
        catch (InvalidOperationException ex)
        {
            logger.Error($"Source configuration error: {ex.Message}");
            return ExitCodes.InvalidArguments;
        }

        // ── 7. Cancellation (Ctrl+C) ────────────────────────────────────────
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            logger.Warn("Cancellation requested (Ctrl+C). Stopping after current file...");
            cts.Cancel();
        };

        // ── 8. Run with outer retry orchestration ───────────────────────────
        var copier      = new FileCopier(source, logger, settings.CopyOptions);
        var orchestrator = new RetryOrchestrator(settings.RetryPolicy, logger);

        int exitCode = await orchestrator.RunAsync(
            ct => copier.RunAsync(folderName, destRoot, filters, ct),
            cts.Token);

        graphHttp?.Dispose();
        return exitCode;
    }

    static void PrintUsage()
    {
        Console.WriteLine("""
            ╔══════════════════════════════════════════════════════════════════╗
            ║  OneDrive / SharePoint / Azure Blob → Network Drive Copier v4.0 ║
            ╚══════════════════════════════════════════════════════════════════╝

            USAGE:
              OneDriveCopier.exe --folder-name <date-folder> [options]

            REQUIRED:
              --folder-name   The date folder to copy, e.g.  20260522

            OPTIONS:
              --auth-mode   Override AzureAd.AuthMode (OneDrive/SharePoint only)
                              ClientCredentials  (default — non-interactive)
                              Interactive        (browser login)
              --dest        Override Destination.RootPath from appsettings.json
              --log-dir     Override log file directory
              --help        Show this help

            SOURCE MODES  (SourceMode in appsettings.json):
              OneDrive    Reads from a user's OneDrive via Graph REST API
              SharePoint  Reads from a SharePoint document library via Graph REST API
              AzureBlob   Reads from an Azure Blob Storage container

            AUTH MODES — OneDrive / SharePoint  (AzureAd.AuthMode):
              ClientCredentials  App-only. Requires AzureAd.ClientSecret.
                                 App needs APPLICATION permissions (admin-consented):
                                   Files.Read.All, Sites.Read.All
              Interactive        Browser pop-up on first run; MSAL token cached to disk.
                                 App needs DELEGATED permissions:
                                   Files.Read, Files.Read.All, Sites.Read.All

            AUTH MODES — AzureBlob  (AzureBlob.AuthMode):
              ConnectionString / ClientCredentials / ManagedIdentity

            RETRY POLICY  (RetryPolicy in appsettings.json):
              MaxAttempts        Max full-run attempts before giving up    (default: 30)
              DelayMinutes       Minutes to wait between attempts          (default: 10)
              MaxDurationHours   Hard time ceiling in hours                (default: 5.0)
              Retries on:  SourceNotFound (folder not yet available), TotalFailure
              No retry on: InvalidArguments, DestUnreachable, PartialFailure, AuthFailure

            EXIT CODES:
              0  Success            4  Partial failure (some files failed — no retry)
              1  Invalid args       5  Total failure   (retried per policy)
              2  Folder not found   6  Auth error      (no retry)
              3  Dest unreachable
            """);
    }
}
