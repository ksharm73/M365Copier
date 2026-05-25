namespace OneDriveCopier;

/// <summary>
/// Core copy engine.
/// Downloads files from the IFileSource (OneDrive or SharePoint via Graph API)
/// and writes them to a local/network destination path.
///
/// Features:
///   - Mirrors the date-folder name at the destination root
///   - Per-file retry with configurable back-off
///   - Post-copy file-size integrity check
///   - Write-probe on destination before touching any files
///   - Returns standardised exit code for Autosys
/// </summary>
internal sealed class FileCopier
{
    private readonly IFileSource         _source;
    private readonly FileLogger          _log;
    private readonly CopyOptionsSettings _opts;

    public FileCopier(IFileSource source, FileLogger log, CopyOptionsSettings opts)
    {
        _source = source;
        _log    = log;
        _opts   = opts;
    }

    public async Task<int> RunAsync(
        string dateFolderName,
        string destRootPath,
        string[] filters,
        CancellationToken ct = default)
    {
        _log.Info($"Source       : {_source.SourceLabel}");
        _log.Info($"Date folder  : {dateFolderName}");
        _log.Info($"Destination  : {destRootPath}");
        _log.Info($"Filter(s)    : {string.Join(", ", filters)}");
        _log.Info($"Overwrite    : {_opts.Overwrite}");
        _log.Separator();

        // ── 1. Validate / create destination ───────────────────────────────
        // Mirror the date folder under the destination root:
        //   destRoot\20260522
        string resolvedDest = Path.Combine(destRootPath, dateFolderName);
        _log.Info($"Resolved dest: {resolvedDest}");

        try { Directory.CreateDirectory(resolvedDest); }
        catch (Exception ex)
        {
            _log.Error($"Cannot create destination folder: {resolvedDest}");
            _log.Error($"Detail: {ex.Message}");
            return ExitCodes.DestUnreachable;
        }

        if (!IsDestinationWritable(resolvedDest))
        {
            _log.Error($"Destination is not writable: {resolvedDest}");
            return ExitCodes.DestUnreachable;
        }

        // ── 2. List remote files via Graph API ──────────────────────────────
        IReadOnlyList<RemoteFile> files;
        try
        {
            _log.Info("Listing remote files via Graph API...");
            files = await _source.ListFilesAsync(dateFolderName, filters, ct);
        }
        catch (FolderNotFoundException ex)
        {
            _log.Error(ex.Message);
            return ExitCodes.SourceNotFound;
        }
        catch (Exception ex)
        {
            _log.Error($"Failed to list remote files: {ex.Message}");
            return ExitCodes.AuthFailure;
        }

        if (files.Count == 0)
        {
            _log.Warn($"No files matched filter(s) in remote folder '{dateFolderName}'.");
            _log.Summary(0, 0, 0, 0);
            return ExitCodes.Success;
        }

        _log.Info($"Files found  : {files.Count}");
        _log.Separator();

        // ── 3. Copy loop ────────────────────────────────────────────────────
        int copied  = 0;
        int skipped = 0;
        int failed  = 0;

        foreach (RemoteFile remoteFile in files)
        {
            if (ct.IsCancellationRequested) break;

            string destFile = Path.Combine(resolvedDest, remoteFile.RelativePath.Replace('/', Path.DirectorySeparatorChar));

            // Ensure sub-directory exists (for nested structures)
            string? subDir = Path.GetDirectoryName(destFile);
            if (!string.IsNullOrEmpty(subDir))
                Directory.CreateDirectory(subDir);

            // Skip if file exists and overwrite is off
            if (!_opts.Overwrite && File.Exists(destFile))
            {
                _log.Warn($"SKIP  {remoteFile.RelativePath}  (already exists at destination)");
                skipped++;
                continue;
            }

            bool success = await DownloadWithRetryAsync(remoteFile, dateFolderName, destFile, ct);
            if (success) copied++;
            else         failed++;
        }

        // ── 4. Summary ──────────────────────────────────────────────────────
        _log.Summary(files.Count, copied, skipped, failed);

        if (failed == 0)               return ExitCodes.Success;
        if (failed > 0 && copied > 0)  return ExitCodes.PartialFailure;
        return ExitCodes.TotalFailure;
    }

    // ── Private helpers ─────────────────────────────────────────────────────

    private async Task<bool> DownloadWithRetryAsync(
        RemoteFile remoteFile,
        string dateFolderName,
        string destFile,
        CancellationToken ct)
    {
        int attempt = 0;

        while (attempt < _opts.MaxRetries)
        {
            attempt++;
            try
            {
                await using Stream remoteStream =
                    await _source.OpenReadAsync(remoteFile, dateFolderName, ct);

                await using var fileStream = new FileStream(
                    destFile, FileMode.Create, FileAccess.Write, FileShare.None,
                    bufferSize: 81920, useAsync: true);

                await remoteStream.CopyToAsync(fileStream, ct);
                await fileStream.FlushAsync(ct);

                // Integrity check: compare declared remote size vs written bytes
                long writtenBytes = fileStream.Length;
                if (remoteFile.SizeBytes > 0 && writtenBytes != remoteFile.SizeBytes)
                    throw new IOException(
                        $"Size mismatch: remote={remoteFile.SizeBytes} bytes, written={writtenBytes} bytes.");

                _log.Success($"COPY  {remoteFile.RelativePath}  ({FormatBytes(writtenBytes)})");
                return true;
            }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                TryDeletePartial(destFile);

                if (attempt < _opts.MaxRetries)
                {
                    int delayMs = _opts.RetryDelaySeconds * 1000 * attempt;
                    _log.Warn($"Attempt {attempt}/{_opts.MaxRetries} failed for " +
                              $"{remoteFile.RelativePath}: {ex.Message}. " +
                              $"Retrying in {delayMs / 1000}s...");
                    await Task.Delay(delayMs, ct);
                }
                else
                {
                    _log.Error($"FAIL  {remoteFile.RelativePath}  " +
                               $"after {_opts.MaxRetries} attempts: {ex.Message}");
                }
            }
        }

        return false;
    }

    private bool IsDestinationWritable(string destFolder)
    {
        string probe = Path.Combine(destFolder, $".write_probe_{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllText(probe, "probe");
            File.Delete(probe);
            return true;
        }
        catch (Exception ex)
        {
            _log.Warn($"Write-probe failed: {ex.Message}");
            return false;
        }
    }

    private static void TryDeletePartial(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch { /* best-effort cleanup */ }
    }

    private static string FormatBytes(long bytes) => bytes switch
    {
        < 1024               => $"{bytes} B",
        < 1024 * 1024        => $"{bytes / 1024.0:F1} KB",
        < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):F1} MB",
        _                    => $"{bytes / (1024.0 * 1024 * 1024):F2} GB"
    };
}
