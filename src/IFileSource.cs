namespace OneDriveCopier;

/// <summary>
/// Represents a remote file entry returned by a file source.
/// </summary>
internal sealed class RemoteFile
{
    /// <summary>File name including extension, e.g. "Q1_Risk.xlsx"</summary>
    public required string Name { get; init; }

    /// <summary>
    /// Path relative to the requested date folder root.
    /// For files directly inside the date folder this equals Name.
    /// Sub-folder files use forward slashes: "subdir/file.xlsx"
    /// </summary>
    public required string RelativePath { get; init; }

    /// <summary>File size in bytes (used for post-copy integrity check).</summary>
    public long SizeBytes { get; init; }
}

/// <summary>
/// Abstraction over a remote file storage location.
/// Implement this interface to add new source types (e.g. Azure Blob).
///
/// Switch between implementations by changing SourceMode in appsettings.json:
///   "SourceMode": "OneDrive"    → OneDriveSource
///   "SourceMode": "SharePoint"  → SharePointSource
/// </summary>
internal interface IFileSource
{
    /// <summary>Human-readable label shown in logs, e.g. "OneDrive" or "SharePoint".</summary>
    string SourceLabel { get; }

    /// <summary>
    /// Lists all files inside the specified date folder that match the given
    /// extension filters.  Returns an empty list if the folder exists but
    /// contains no matching files.  Throws <see cref="FolderNotFoundException"/>
    /// when the remote folder does not exist.
    /// </summary>
    Task<IReadOnlyList<RemoteFile>> ListFilesAsync(
        string dateFolderName,
        string[] filters,
        CancellationToken ct = default);

    /// <summary>
    /// Opens a read stream for the given remote file.
    /// The caller is responsible for disposing the stream.
    /// </summary>
    Task<Stream> OpenReadAsync(
        RemoteFile file,
        string dateFolderName,
        CancellationToken ct = default);
}

/// <summary>Thrown when the requested remote folder does not exist.</summary>
internal sealed class FolderNotFoundException : Exception
{
    public FolderNotFoundException(string message) : base(message) { }
}
