using System.Net;
using System.Text.Json.Nodes;

namespace OneDriveCopier;

/// <summary>
/// IFileSource implementation — SharePoint document library via raw Graph REST API (no SDK).
///
/// REST endpoints used:
///
///   Step 1 — Resolve site GUID (cached):
///     GET /sites/{hostname}:/{sitePath}
///     e.g. GET /sites/contoso.sharepoint.com:/sites/FinanceTeam
///     Returns { id: "hostname,siteGuid,webGuid", ... }
///
///   Step 2 — Resolve drive ID (cached):
///     GET /sites/{siteId}/drives?$select=id,name
///     Filter client-side by LibraryName.
///
///   List files in date folder:
///     GET /drives/{driveId}/items/root:/{basePath}/{dateFolder}:/children
///         ?$select=name,size,file,folder
///
///   Download file content:
///     GET /drives/{driveId}/items/root:/{basePath}/{dateFolder}/{file}:/content
///
/// Building the URL as a plain string and calling it directly via HttpClient
/// avoids every Graph SDK indexer/OData encoding issue we encountered.
/// </summary>
internal sealed class SharePointSource : IFileSource
{
    private readonly GraphHttpClient    _graph;
    private readonly SharePointSettings _cfg;
    private string?                     _cachedSiteId;
    private string?                     _cachedDriveId;

    public string SourceLabel =>
        $"SharePoint ({_cfg.Hostname}/{_cfg.SitePath}, library: {_cfg.LibraryName})";

    public SharePointSource(GraphHttpClient graph, SharePointSettings cfg)
    {
        _graph = graph;
        _cfg   = cfg;

        if (string.IsNullOrWhiteSpace(cfg.Hostname))
            throw new InvalidOperationException("SharePoint.Hostname is not configured.");
        if (string.IsNullOrWhiteSpace(cfg.SitePath))
            throw new InvalidOperationException("SharePoint.SitePath is not configured.");
        if (string.IsNullOrWhiteSpace(cfg.LibraryName))
            throw new InvalidOperationException("SharePoint.LibraryName is not configured.");
        if (string.IsNullOrWhiteSpace(cfg.BaseFolderPath))
            throw new InvalidOperationException("SharePoint.BaseFolderPath is not configured.");
    }

    // ── IFileSource ──────────────────────────────────────────────────────────

    public async Task<IReadOnlyList<RemoteFile>> ListFilesAsync(
        string dateFolderName,
        string[] filters,
        CancellationToken ct = default)
    {
        string driveId  = await ResolveDriveIdAsync(ct);
        string itemPath = BuildItemPath(dateFolderName);
        string encodedPath = Uri.EscapeDataString(itemPath);
        string url = $"drives/{driveId}/items/root:/{encodedPath}:/children" +
                     "?$select=name,size,file,folder";

        JsonObject page;
        try
        {
            page = await _graph.GetJsonAsync(url, ct);
        }
        catch (HttpRequestException ex)
            when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            throw new FolderNotFoundException(
                $"SharePoint folder not found: /{itemPath} " +
                $"(Site: {_cfg.Hostname}/{_cfg.SitePath}, Library: {_cfg.LibraryName})");
        }

        return await CollectAllPagesAsync(page, filters, ct);
    }

    public async Task<Stream> OpenReadAsync(
        RemoteFile file,
        string dateFolderName,
        CancellationToken ct = default)
    {
        string driveId  = await ResolveDriveIdAsync(ct);
        string itemPath = BuildItemPath(dateFolderName) + "/" +
                          file.RelativePath.Replace('\\', '/');
        string encodedPath = Uri.EscapeDataString(itemPath);
        string url = $"drives/{driveId}/items/root:/{encodedPath}:/content";

        return await _graph.GetStreamAsync(url, ct);
    }

    // ── Private ──────────────────────────────────────────────────────────────

    private async Task<string> ResolveSiteIdAsync(CancellationToken ct)
    {
        if (_cachedSiteId is not null) return _cachedSiteId;

        string hostname = _cfg.Hostname.Trim().TrimEnd('/');
        string sitePath = _cfg.SitePath.Trim().Trim('/');   // e.g. "sites/FinanceTeam"

        // Direct path-based site lookup — plain string URL, no SDK encoding issues.
        // Graph REST: GET /sites/{hostname}:/{sitePath}
        string url = $"sites/{hostname}:/{sitePath}";

        JsonObject site;
        try
        {
            site = await _graph.GetJsonAsync(url, ct);
        }
        catch (HttpRequestException ex)
            when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            throw new InvalidOperationException(
                $"SharePoint site not found: {hostname}/{sitePath}. " +
                "Verify SharePoint.Hostname and SharePoint.SitePath in appsettings.json.");
        }

        _cachedSiteId = site["id"]?.GetValue<string>()
            ?? throw new InvalidOperationException(
                $"Graph returned no site id for: {hostname}/{sitePath}");

        return _cachedSiteId;
    }

    private async Task<string> ResolveDriveIdAsync(CancellationToken ct)
    {
        if (_cachedDriveId is not null) return _cachedDriveId;

        string siteId = await ResolveSiteIdAsync(ct);

        // siteId is the compound GUID Graph returns: "hostname,siteGuid,webGuid"
        // URL-encode it so the commas don't break the path segment.
        string encodedSiteId = Uri.EscapeDataString(siteId);
        string url = $"sites/{encodedSiteId}/drives?$select=id,name";

        JsonObject drivesPage = await _graph.GetJsonAsync(url, ct);

        if (drivesPage["value"] is not JsonArray driveList)
            throw new InvalidOperationException(
                $"Graph returned no drives array for site: {siteId}");

        // Find the library by display name (case-insensitive)
        JsonObject? match = driveList
            .OfType<JsonObject>()
            .FirstOrDefault(d =>
                string.Equals(
                    d["name"]?.GetValue<string>(),
                    _cfg.LibraryName,
                    StringComparison.OrdinalIgnoreCase));

        if (match is null)
        {
            var available = string.Join(", ",
                driveList.OfType<JsonObject>()
                         .Select(d => d["name"]?.GetValue<string>()));
            throw new InvalidOperationException(
                $"Document library '{_cfg.LibraryName}' not found in site " +
                $"'{_cfg.Hostname}/{_cfg.SitePath}'. " +
                $"Available libraries: [{available}]");
        }

        _cachedDriveId = match["id"]?.GetValue<string>()
            ?? throw new InvalidOperationException(
                $"Graph returned no id for library '{_cfg.LibraryName}'");

        return _cachedDriveId;
    }

    private string BuildItemPath(string dateFolderName)
    {
        string base_ = _cfg.BaseFolderPath.Trim('/').Trim('\\');
        return $"{base_}/{dateFolderName}";
    }

    private async Task<IReadOnlyList<RemoteFile>> CollectAllPagesAsync(
        JsonObject firstPage,
        string[] filters,
        CancellationToken ct)
    {
        var result = new List<RemoteFile>();
        JsonObject? page = firstPage;

        while (page is not null)
        {
            if (page["value"] is JsonArray items)
            {
                foreach (JsonNode? node in items)
                {
                    if (node is not JsonObject item) continue;
                    if (item["file"] is null) continue;  // skip folders

                    string? name = item["name"]?.GetValue<string>();
                    if (name is null) continue;
                    if (!MatchesFilter(name, filters)) continue;

                    result.Add(new RemoteFile
                    {
                        Name         = name,
                        RelativePath = name,
                        SizeBytes    = item["size"]?.GetValue<long>() ?? 0
                    });
                }
            }

            string? nextLink = page["@odata.nextLink"]?.GetValue<string>();
            page = nextLink is not null
                ? await _graph.GetJsonByFullUrlAsync(nextLink, ct)
                : null;
        }

        return result;
    }

    private static bool MatchesFilter(string fileName, string[] filters)
    {
        if (filters.Length == 1 && filters[0] == "*.*") return true;
        string ext = Path.GetExtension(fileName).ToLowerInvariant();
        return filters.Any(f =>
        {
            string fe = f.StartsWith("*.") ? f[1..].ToLowerInvariant() : f.ToLowerInvariant();
            return ext == fe || f == "*.*";
        });
    }
}
