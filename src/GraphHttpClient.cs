using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Identity.Client;

namespace OneDriveCopier;

/// <summary>
/// Replaces the Graph SDK entirely.
/// Handles OAuth 2.0 token acquisition and wraps HttpClient with:
///   • Automatic Bearer token attachment
///   • Token refresh when access token expires
///   • Two auth modes:
///       ClientCredentials — raw POST to /token endpoint (no MSAL needed,
///                           but MSAL.NET is used here for consistency and
///                           because it handles expiry correctly)
///       Interactive       — MSAL.NET InteractiveAuthProvider with file-based
///                           token cache (browser only on first run)
///
/// All Graph REST calls go through SendAsync() or the typed helpers
/// GetJsonAsync() / GetStreamAsync().
/// </summary>
internal sealed class GraphHttpClient : IDisposable
{
    private const string GraphBase  = "https://graph.microsoft.com/v1.0";
    private const string GraphScope = "https://graph.microsoft.com/.default";

    private static readonly string[] DelegatedScopes =
    [
        "https://graph.microsoft.com/Files.Read",
        "https://graph.microsoft.com/Files.Read.All",
        "https://graph.microsoft.com/Sites.Read.All"
    ];

    private readonly HttpClient        _http;
    private readonly IConfidentialClientApplication? _ccApp;   // client credentials
    private readonly IPublicClientApplication?        _pcaApp;  // interactive
    private readonly bool              _isInteractive;
    private readonly FileLogger        _log;

    // ── Factory ──────────────────────────────────────────────────────────────

    public static GraphHttpClient Create(
        AzureAdSettings cfg,
        string? authModeOverride,
        FileLogger log)
    {
        if (string.IsNullOrWhiteSpace(cfg.TenantId))
            throw new InvalidOperationException("AzureAd.TenantId is not configured.");
        if (string.IsNullOrWhiteSpace(cfg.ClientId))
            throw new InvalidOperationException("AzureAd.ClientId is not configured.");

        string mode = (authModeOverride ?? cfg.AuthMode).Trim();

        return mode.ToLowerInvariant() switch
        {
            "clientcredentials" => CreateClientCredentials(cfg, log),
            "interactive"       => CreateInteractive(cfg, log),
            _ => throw new InvalidOperationException(
                $"Unknown AuthMode '{mode}'. Valid values: ClientCredentials, Interactive")
        };
    }

    // ── HTTP helpers ─────────────────────────────────────────────────────────

    /// <summary>
    /// GET {GraphBase}/{relativeUrl}  and deserialize the JSON response body.
    /// relativeUrl should NOT start with a slash.
    /// </summary>
    public async Task<JsonObject> GetJsonAsync(string relativeUrl, CancellationToken ct)
    {
        string token = await AcquireTokenAsync(ct);
        using var req = new HttpRequestMessage(
            HttpMethod.Get, $"{GraphBase}/{relativeUrl}");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        using HttpResponseMessage resp = await _http.SendAsync(req, ct);
        await EnsureSuccessAsync(resp, relativeUrl);

        string body = await resp.Content.ReadAsStringAsync(ct);
        return JsonNode.Parse(body)?.AsObject()
            ?? throw new InvalidOperationException(
                $"Graph returned empty body for: {relativeUrl}");
    }

    /// <summary>
    /// GET a full URL (used for @odata.nextLink pagination and download URLs).
    /// </summary>
    public async Task<JsonObject> GetJsonByFullUrlAsync(string url, CancellationToken ct)
    {
        string token = await AcquireTokenAsync(ct);
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        using HttpResponseMessage resp = await _http.SendAsync(req, ct);
        await EnsureSuccessAsync(resp, url);

        string body = await resp.Content.ReadAsStringAsync(ct);
        return JsonNode.Parse(body)?.AsObject()
            ?? throw new InvalidOperationException($"Graph returned empty body for: {url}");
    }

    /// <summary>
    /// GET a file download stream via a Graph content URL.
    /// Graph returns a 302 redirect to a pre-authenticated CDN URL.
    /// HttpClient follows redirects automatically (AllowAutoRedirect = true).
    /// </summary>
    public async Task<Stream> GetStreamAsync(string relativeUrl, CancellationToken ct)
    {
        string token = await AcquireTokenAsync(ct);
        using var req = new HttpRequestMessage(
            HttpMethod.Get, $"{GraphBase}/{relativeUrl}");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        // Do NOT dispose the response here — the caller reads the stream
        HttpResponseMessage resp = await _http.SendAsync(
            req, HttpCompletionOption.ResponseHeadersRead, ct);
        await EnsureSuccessAsync(resp, relativeUrl);

        return await resp.Content.ReadAsStreamAsync(ct);
    }

    public void Dispose() => _http.Dispose();

    // ── Private constructors ─────────────────────────────────────────────────

    private GraphHttpClient(
        IConfidentialClientApplication ccApp,
        FileLogger log)
    {
        _ccApp         = ccApp;
        _isInteractive = false;
        _log           = log;
        _http          = BuildHttpClient();
    }

    private GraphHttpClient(
        IPublicClientApplication pcaApp,
        FileLogger log)
    {
        _pcaApp        = pcaApp;
        _isInteractive = true;
        _log           = log;
        _http          = BuildHttpClient();
    }

    private static GraphHttpClient CreateClientCredentials(
        AzureAdSettings cfg, FileLogger log)
    {
        if (string.IsNullOrWhiteSpace(cfg.ClientSecret))
            throw new InvalidOperationException(
                "AzureAd.ClientSecret is required when AuthMode = \"ClientCredentials\".");

        log.Info("Auth mode    : ClientCredentials (app-only, non-interactive)");

        // MSAL ConfidentialClientApplication handles token caching + expiry
        // transparently. Under the hood it POSTs to:
        //   https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
        var app = ConfidentialClientApplicationBuilder
            .Create(cfg.ClientId)
            .WithTenantId(cfg.TenantId)
            .WithClientSecret(cfg.ClientSecret)
            .Build();

        return new GraphHttpClient(app, log);
    }

    private static GraphHttpClient CreateInteractive(
        AzureAdSettings cfg, FileLogger log)
    {
        log.Info("Auth mode    : Interactive (browser login)");

        string cacheDir = string.IsNullOrWhiteSpace(cfg.InteractiveLogin.TokenCacheDir)
            ? AppContext.BaseDirectory
            : cfg.InteractiveLogin.TokenCacheDir;

        string cachePath = Path.Combine(cacheDir, "onedrive_copier_token.bin");
        log.Info($"Token cache  : {cachePath}");
        log.Info("A browser window will open if no valid cached token is found.");
        log.Info($"Redirect URI : {cfg.InteractiveLogin.RedirectUri}");

        var app = PublicClientApplicationBuilder
            .Create(cfg.ClientId)
            .WithTenantId(cfg.TenantId)
            .WithRedirectUri(cfg.InteractiveLogin.RedirectUri)
            .Build();

        // Wire up file-based token cache (Windows DPAPI encrypted)
        RegisterFileTokenCache(app.UserTokenCache, cachePath);

        return new GraphHttpClient(app, log);
    }

    // ── Token acquisition ────────────────────────────────────────────────────

    private async Task<string> AcquireTokenAsync(CancellationToken ct)
    {
        if (!_isInteractive)
        {
            // Client credentials — always hits the in-memory MSAL cache first,
            // only calls /token when the cached token has expired.
            var result = await _ccApp!
                .AcquireTokenForClient([GraphScope])
                .ExecuteAsync(ct);
            return result.AccessToken;
        }
        else
        {
            // Interactive — try silent (cached) first; fall back to browser.
            var accounts = await _pcaApp!.GetAccountsAsync();
            try
            {
                var silent = await _pcaApp
                    .AcquireTokenSilent(DelegatedScopes, accounts.FirstOrDefault())
                    .ExecuteAsync(ct);
                return silent.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                _log.Info("No cached token found — opening browser for sign-in...");
                var interactive = await _pcaApp
                    .AcquireTokenInteractive(DelegatedScopes)
                    .ExecuteAsync(ct);
                return interactive.AccessToken;
            }
        }
    }

    // ── File-based MSAL token cache ──────────────────────────────────────────

    private static void RegisterFileTokenCache(ITokenCache cache, string path)
    {
        // Simple file-based persistence. On Windows, ProtectedData provides
        // DPAPI encryption so the token file is user-scoped.
        cache.SetBeforeAccess(args =>
        {
            if (File.Exists(path))
            {
                try
                {
                    byte[] encrypted = File.ReadAllBytes(path);
                    byte[] plain     = System.Security.Cryptography.ProtectedData.Unprotect(
                        encrypted, null,
                        System.Security.Cryptography.DataProtectionScope.CurrentUser);
                    args.TokenCache.DeserializeMsalV3(plain);
                }
                catch { /* corrupted cache — start fresh */ }
            }
        });

        cache.SetAfterAccess(args =>
        {
            if (!args.HasStateChanged) return;
            try
            {
                byte[] plain     = args.TokenCache.SerializeMsalV3();
                byte[] encrypted = System.Security.Cryptography.ProtectedData.Protect(
                    plain, null,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.WriteAllBytes(path, encrypted);
            }
            catch { /* best-effort: next run will re-authenticate */ }
        });
    }

    // ── Shared ───────────────────────────────────────────────────────────────

    private static HttpClient BuildHttpClient()
    {
        var handler = new HttpClientHandler { AllowAutoRedirect = true };
        var client  = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(120)
        };
        client.DefaultRequestHeaders.Accept
              .Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage resp, string url)
    {
        if (resp.IsSuccessStatusCode) return;

        string body = await resp.Content.ReadAsStringAsync();
        string detail;
        try
        {
            var obj = JsonNode.Parse(body);
            detail  = obj?["error"]?["message"]?.GetValue<string>()
                   ?? obj?["error"]?.GetValue<string>()
                   ?? body;
        }
        catch { detail = body; }

        throw new HttpRequestException(
            $"Graph API {(int)resp.StatusCode} for [{url}]: {detail}",
            inner: null,
            statusCode: resp.StatusCode);
    }
}
