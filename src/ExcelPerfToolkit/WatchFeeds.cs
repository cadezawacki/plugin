using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPerfToolkit;

/// <summary>
/// Filesystem-change RTD feeds. These plug into the same <see cref="Feed"/> framework as the
/// demo feeds in <c>RtdServer.cs</c>, but instead of a polling loop they arm a
/// <see cref="FileSystemWatcher"/> and publish a monotonically increasing change counter as
/// their latest value. The RTD server's 250 ms throttle naturally coalesces bursts of
/// filesystem events into a single push, so a cell watching a file updates at most a few
/// times a second no matter how chatty the underlying events are.
///
/// <para>Feed specs (created by <c>FeedManager.CreateFeed</c>): <c>watchfile:&lt;path&gt;</c>
/// and <c>watchfolder:&lt;path&gt;</c>. Paths must be fully qualified - a relative path would
/// silently resolve against Excel's drifting process directory. The counter starts at 0 and
/// increments on every create/change/delete/rename event, so any dependent formula
/// recalculates when the watched target changes.</para>
/// </summary>
internal sealed class WatchFileFeed : Feed
{
    private readonly string _directory;
    private readonly string _filter;
    private long _count;

    public WatchFileFeed(string spec, string path) : base(spec)
    {
        if (!Path.IsPathFullyQualified(path))
        {
            throw new ArgumentException(
                $"EPT.WATCHFILE requires an absolute path; got '{path}'. A relative path would resolve against Excel's process directory, which drifts at runtime.");
        }
        var dir = Path.GetDirectoryName(path);
        _directory = string.IsNullOrEmpty(dir) ? Path.GetPathRoot(path) ?? path : dir;
        var name = Path.GetFileName(path);
        _filter = string.IsNullOrEmpty(name) ? "*.*" : name;
        LatestValue = (double)0;
    }

    protected override Task RunAsync(CancellationToken cancellationToken)
        => WatchFeedRunner.RunAsync(_directory, _filter, includeSubdirectories: false, Bump, cancellationToken);

    private void Bump() => LatestValue = (double)Interlocked.Increment(ref _count);
}

/// <summary>
/// Watches an entire folder (non-recursive) for create/change/delete/rename events and
/// publishes a change counter. See <see cref="WatchFileFeed"/> for the throttling notes.
/// </summary>
internal sealed class WatchFolderFeed : Feed
{
    private readonly string _directory;
    private long _count;

    public WatchFolderFeed(string spec, string path) : base(spec)
    {
        if (!Path.IsPathFullyQualified(path))
        {
            throw new ArgumentException(
                $"EPT.WATCHFOLDER requires an absolute path; got '{path}'. A relative path would resolve against Excel's process directory, which drifts at runtime.");
        }
        _directory = path;
        LatestValue = (double)0;
    }

    protected override Task RunAsync(CancellationToken cancellationToken)
        => WatchFeedRunner.RunAsync(_directory, "*.*", includeSubdirectories: false, Bump, cancellationToken);

    private void Bump() => LatestValue = (double)Interlocked.Increment(ref _count);
}

/// <summary>
/// Shared driver: arms a <see cref="FileSystemWatcher"/>, routes every event to the change
/// callback, and parks until the watcher dies or the feed is cancelled. The first arm failure
/// (e.g. a directory that does not exist) propagates, so the subscribed cell shows the
/// documented <c>#N/A</c>. After a successful arm, a FATAL watcher error - watched directory
/// deleted or renamed, network share dropped, handle invalidated - no longer leaves a
/// silently dead watcher behind: the driver bumps once (events may have been missed), then
/// re-arms with a short backoff until the directory is watchable again or the feed stops.
/// <see cref="InternalBufferOverflowException"/> is recoverable (the watcher keeps running)
/// and only triggers a bump, which is exactly the right coalesced signal for a change counter.
/// </summary>
internal static class WatchFeedRunner
{
    /// <summary>Backoff between re-arm attempts after the watch goes down.</summary>
    private static readonly TimeSpan RearmDelay = TimeSpan.FromSeconds(2);

    public static async Task RunAsync(string directory, string filter, bool includeSubdirectories, Action onChange, CancellationToken cancellationToken)
    {
        var firstArm = true;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var armed = false;
            try
            {
                await RunOneWatcherAsync(directory, filter, includeSubdirectories, onChange, !firstArm, cancellationToken).ConfigureAwait(false);
                // Returning (rather than throwing) means the watcher armed and later
                // died from a fatal error.
                armed = true;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception) when (!firstArm)
            {
                // Re-arm attempt failed (directory still missing / share still down):
                // retry quietly below - bumping here would recalc dependents every
                // backoff tick for as long as the target is absent.
            }
            firstArm = false;
            if (armed)
            {
                // The live watch tore down; events may have been missed in the gap.
                onChange();
            }
            await Task.Delay(RearmDelay, cancellationToken).ConfigureAwait(false);
        }
    }

    private static async Task RunOneWatcherAsync(string directory, string filter, bool includeSubdirectories, Action onChange, bool bumpOnArm, CancellationToken cancellationToken)
    {
        using var watcher = new FileSystemWatcher(directory, filter)
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size | NotifyFilters.CreationTime,
            IncludeSubdirectories = includeSubdirectories,
        };

        var fatal = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        void OnChanged(object? sender, FileSystemEventArgs e) => onChange();
        void OnRenamed(object? sender, RenamedEventArgs e) => onChange();
        void OnError(object? sender, ErrorEventArgs e)
        {
            onChange();
            if (e.GetException() is not InternalBufferOverflowException)
            {
                fatal.TrySetResult();
            }
        }

        watcher.Created += OnChanged;
        watcher.Changed += OnChanged;
        watcher.Deleted += OnChanged;
        watcher.Renamed += OnRenamed;
        watcher.Error += OnError;
        try
        {
            watcher.EnableRaisingEvents = true;
            if (bumpOnArm)
            {
                // Regained the watch after a gap: signal dependents that the target may
                // have changed unobserved.
                onChange();
            }
            await fatal.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            watcher.EnableRaisingEvents = false;
            watcher.Created -= OnChanged;
            watcher.Changed -= OnChanged;
            watcher.Deleted -= OnChanged;
            watcher.Renamed -= OnRenamed;
            watcher.Error -= OnError;
        }
    }
}
