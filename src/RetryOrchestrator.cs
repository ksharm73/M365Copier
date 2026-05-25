namespace OneDriveCopier;

/// <summary>
/// Outer retry orchestrator — distinct from the per-file copy retries in FileCopier.
///
/// Purpose: an Autosys job that runs every 10 minutes for up to 5 hours waiting
/// for the source folder to become available (e.g. upstream ETL finishes at an
/// unpredictable time). All three values are configurable in appsettings.json:
///
///   RetryPolicy.MaxAttempts       — max number of full run attempts  (default: 30)
///   RetryPolicy.DelayMinutes      — wait between attempts in minutes (default: 10)
///   RetryPolicy.MaxDurationHours  — hard time ceiling in hours       (default: 5.0)
///
/// Retry triggers:
///   ExitCodes.SourceNotFound (2) — folder doesn't exist yet; worth retrying
///   ExitCodes.TotalFailure   (5) — all files failed; may be transient
///
/// No retry on:
///   ExitCodes.Success          (0) — done
///   ExitCodes.InvalidArguments (1) — config problem; retrying won't help
///   ExitCodes.DestUnreachable  (3) — network drive issue; operator action needed
///   ExitCodes.PartialFailure   (4) — some files copied; operator should review
///   ExitCodes.AuthFailure      (6) — credentials problem; retrying won't help
/// </summary>
internal sealed class RetryOrchestrator
{
    private readonly RetryPolicySettings _policy;
    private readonly FileLogger          _log;

    public RetryOrchestrator(RetryPolicySettings policy, FileLogger log)
    {
        _policy = policy;
        _log    = log;
    }

    /// <summary>
    /// Runs <paramref name="operation"/> repeatedly according to the retry policy.
    /// Returns the final exit code.
    /// </summary>
    public async Task<int> RunAsync(
        Func<CancellationToken, Task<int>> operation,
        CancellationToken ct = default)
    {
        DateTime deadline = DateTime.UtcNow.AddHours(_policy.MaxDurationHours);
        int attempt = 0;
        int lastExitCode = ExitCodes.TotalFailure;

        _log.Info($"Retry policy : max {_policy.MaxAttempts} attempts, " +
                  $"{_policy.DelayMinutes} min delay, " +
                  $"{_policy.MaxDurationHours}h ceiling");

        while (attempt < _policy.MaxAttempts && DateTime.UtcNow < deadline)
        {
            if (ct.IsCancellationRequested) break;

            attempt++;
            bool isFirstAttempt = attempt == 1;

            if (!isFirstAttempt)
            {
                TimeSpan remaining = deadline - DateTime.UtcNow;
                _log.Info($"Attempt {attempt}/{_policy.MaxAttempts} " +
                          $"(time remaining: {remaining:hh\\:mm\\:ss})");
            }
            else
            {
                _log.Info($"Attempt {attempt}/{_policy.MaxAttempts}");
            }

            _log.Separator();

            try
            {
                lastExitCode = await operation(ct);
            }
            catch (OperationCanceledException)
            {
                _log.Warn("Operation cancelled.");
                return ExitCodes.TotalFailure;
            }
            catch (Exception ex)
            {
                _log.Error($"Unhandled exception on attempt {attempt}: {ex.Message}");
                lastExitCode = ExitCodes.TotalFailure;
            }

            // Decide whether to retry
            if (!ShouldRetry(lastExitCode))
            {
                if (lastExitCode == ExitCodes.Success)
                    _log.Success($"Completed successfully on attempt {attempt}.");
                else
                    _log.Info($"Exit code {lastExitCode} — no retry warranted.");
                return lastExitCode;
            }

            // Check if we have room for another attempt
            bool attemptsExhausted = attempt >= _policy.MaxAttempts;
            bool deadlineReached   = DateTime.UtcNow.AddMinutes(_policy.DelayMinutes) >= deadline;

            if (attemptsExhausted || deadlineReached)
            {
                string reason = attemptsExhausted
                    ? $"max attempts ({_policy.MaxAttempts}) reached"
                    : $"time ceiling ({_policy.MaxDurationHours}h) would be exceeded";
                _log.Error($"Giving up after attempt {attempt} — {reason}. " +
                           $"Last exit code: {lastExitCode}.");
                return lastExitCode;
            }

            // Wait before next attempt
            TimeSpan delay = TimeSpan.FromMinutes(_policy.DelayMinutes);
            DateTime nextAttemptAt = DateTime.Now.Add(delay);
            _log.Warn($"Attempt {attempt} exit code {lastExitCode} — retrying. " +
                      $"Next attempt at {nextAttemptAt:HH:mm:ss} " +
                      $"(waiting {_policy.DelayMinutes} min)...");

            try
            {
                await Task.Delay(delay, ct);
            }
            catch (OperationCanceledException)
            {
                _log.Warn("Wait interrupted by cancellation.");
                return lastExitCode;
            }
        }

        return lastExitCode;
    }

    private static bool ShouldRetry(int exitCode) => exitCode is
        ExitCodes.SourceNotFound or  // folder not yet available — normal during ETL wait
        ExitCodes.TotalFailure;      // all files failed — may be transient
}
