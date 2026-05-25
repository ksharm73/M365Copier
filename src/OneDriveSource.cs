using System.Net;
using System.Text.Json.Nodes;

namespace OneDriveCopier;

/// <summary>
/// IFileSource implementation — OneDrive via raw Graph REST API (no SDK).
///
/// REST endpoints used:
///
///   Resolve drive ID (cached once per run):
///     GET /me/drive                      → { id }   (interactive, no UserId)
///     GET /users/{userId}/drive          → { id }   (app-only or interactive with UserId)
///
///   List files in date folder:
///     GET /drives/{driveId}/items/root:/{basePath}/{dateFolder}:/children
///         ?$select=name,size,file,folder
///
///   Download file content:
///     GET /drives/{driveId}/items/root:/{basePath}/{dateFolder}/{file}:/content
///         (Graph returns 302 → CDN URL; HttpClient follows redirect automatically)
/// </summary>
internal sealed class OneDriveSource : IFileSource
{
    private readonly GraphHttpClient  _graph;
    private readonly OneDriveSettings _cfg;
    private readonly bool             _useMeDrive;
    private string?                   _cachedDriveId;

    public string SourceLabel => _useMeDrive
        ? "OneDrive (signed-in user /me)"
        : $"OneDrive (user: {_cfg.UserId})";

    public OneDriveSource(
        GraphHttpClient graph,
        OneDriveSettings cfg,
        bool isInteractiveAuth)
    {
        _graph = graph;
        _cfg   = cfg;

        if (string.IsNullOrWhiteSpace(cfg.BaseFolderPath))
            throw new InvalidOperationException(
                "OneDrive.BaseFolderPath is not configured in appsettings.json.");

        if (!isInteractiveAuth && string.IsNullOrWhiteSpace(cfg.UserId))
            throw new InvalidOperationException(
                "OneDrive.UserId is required when AuthMode = \"ClientCredentials\". " +
                "Set it to the UPN or object ID of the drive owner.");

        _useMeDrive = isInteractiveAuth && string.IsNullOrWhiteSpace(cfg.UserId);
    }

    // ── IFileSource ──────────────────────────────────────────────────────────

    public async Task<IReadOnlyList<RemoteFile>> ListFilesAsync(
        string dateFolderName,
        string[] filters,
        CancellationToken ct = default)
    {
        string driveId  = await ResolveDriveIdAsync(ct);
        string itemPath = BuildItemPath(dateFolderName);

        // Encode the path for the Graph path-item URL segment
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
            string who = _useMeDrive ? "signed-in user" : _cfg.UserId;
            throw new FolderNotFoundException(
                $"OneDrive folder not found: /{itemPath} " +
                $"(Drive owner: {who}, BasePath: {_cfg.BaseFolderPath})");
        }

        return await CollectAllPagesAsync(page, filters, ct);
    }

    public async Task<Stream> OpenReadAsync(
        RemoteFile file,
        string dateFolderName,
        CancellationToken ct = default)
    {
        string driveId   = await ResolveDriveIdAsync(ct);
        string itemPath  = BuildItemPath(dateFolderName) + "/" +
                           file.RelativePath.Replace('\\', '/');
        string encodedPath = Uri.EscapeDataString(itemPath);
        string url = $"drives/{driveId}/items/root:/{encodedPath}:/content";

        return await _graph.GetStreamAsync(url, ct);
    }

    // ── Private ──────────────────────────────────────────────────────────────

    private async Task<string> ResolveDriveIdAsync(CancellationToken ct)
    {
        if (_cachedDriveId is not null) return _cachedDriveId;

        string url = _useMeDrive ? "me/drive" : $"users/{_cfg.UserId}/drive";

        JsonObject drive = await _graph.GetJsonAsync(url, ct);

        _cachedDriveId = drive["id"]?.GetValue<string>()
            ?? throw new InvalidOperationException(
                $"Graph returned no drive id for: {url}. " +
                "Verify the UserId and that the account has a OneDrive licence.");

        return _cachedDriveId;
    }

    private string BuildItemPath(string dateFolderName)
    {
        string base_ = _cfg.BaseFolderPath.Trim('/').Trim('\\');
        return $"{base_}/{dateFolderName}";
    }

    /// <summary>
    /// Iterates all @odata.nextLink pages and collects matching files.
    /// Graph returns max 200 items per page for /children.
    /// </summary>
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

                    // Skip folders — only process file items
                    if (item["file"] is null) continue;

                    string? name = item["name"]?.GetValue<string>();
                    if (name is null) continue;
                    if (!MatchesFilter(name, filters)) continue;

                    long size = item["size"]?.GetValue<long>() ?? 0;

                    result.Add(new RemoteFile
                    {
                        Name         = name,
                        RelativePath = name,
                        SizeBytes    = size
                    });
                }
            }

            // Follow next page if present
            string? nextLink = page["@odata.nextLink"]?.GetValue<string>();
            if (nextLink is not null)
                page = await _graph.GetJsonByFullUrlAsync(nextLink, ct);
            else
                page = null;
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
