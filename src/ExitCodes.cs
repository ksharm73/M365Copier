namespace OneDriveCopier;

/// <summary>
/// Standardised exit codes for Autosys exit_code_range conditions.
/// </summary>
internal static class ExitCodes
{
    public const int Success          = 0;  // All files downloaded
    public const int InvalidArguments = 1;  // Bad CLI args or missing config values
    public const int SourceNotFound   = 2;  // Remote folder not found via Graph API
    public const int DestUnreachable  = 3;  // Network drive not writable
    public const int PartialFailure   = 4;  // Some files failed
    public const int TotalFailure     = 5;  // No files downloaded
    public const int AuthFailure      = 6;  // Graph API authentication error
}
