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
/// and <c>watchfolder:&lt;path&gt;</c>. The counter starts at 0 and increments on every
/// create/change/delete/rename event, so any dependent formula recalculates when the watched
/// target changes.</para>
/// </summary>
internal sealed class WatchFileFeed : Feed
{
    private readonly string _directory;
    private readonly string _filter;
    private long _count;

    public WatchFileFeed(string spec, string path) : base(spec)
    {
        var dir = Path.GetDirectoryName(path);
        _directory = string.IsNullOrEmpty(dir) ? "." : dir;
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
        _directory = string.IsNullOrEmpty(path) ? "." : path;
        LatestValue = (double)0;
    }

    protected override Task RunAsync(CancellationToken cancellationToken)
        => WatchFeedRunner.RunAsync(_directory, "*.*", includeSubdirectories: false, Bump, cancellationToken);

    private void Bump() => LatestValue = (double)Interlocked.Increment(ref _count);
}

/// <summary>
/// Shared driver: arms a <see cref="FileSystemWatcher"/>, routes every event to
/// <paramref name="onChange"/>, and parks until cancellation. The watcher is disposed and its
/// handlers detached on exit. A non-existent directory makes the constructor throw, which the
/// <see cref="Feed"/> base class surfaces as <c>#N/A</c> in the subscribed cell.
/// </summary>
internal static class WatchFeedRunner
{
    public static async Task RunAsync(string directory, string filter, bool includeSubdirectories, Action onChange, CancellationToken cancellationToken)
    {
        using var watcher = new FileSystemWatcher(directory, filter)
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size | NotifyFilters.CreationTime,
            IncludeSubdirectories = includeSubdirectories,
        };

        void OnChanged(object? sender, FileSystemEventArgs e) => onChange();
        void OnRenamed(object? sender, RenamedEventArgs e) => onChange();
        void OnError(object? sender, ErrorEventArgs e) => onChange();

        watcher.Created += OnChanged;
        watcher.Changed += OnChanged;
        watcher.Deleted += OnChanged;
        watcher.Renamed += OnRenamed;
        watcher.Error += OnError;
        watcher.EnableRaisingEvents = true;
        try
        {
            await Task.Delay(Timeout.Infinite, cancellationToken).ConfigureAwait(false);
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
