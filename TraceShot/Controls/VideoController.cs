using System;
using System.Windows.Controls;
using System.Windows.Threading;

namespace TraceShot.Controls
{
    public class VideoPositionEventArgs : EventArgs
    {
        public TimeSpan Current { get; }
        public TimeSpan Total { get; }
        public bool IsPlaying { get; }

        public VideoPositionEventArgs(TimeSpan current, TimeSpan total, bool isPlaying)
        {
            Current = current; Total = total; IsPlaying = isPlaying;
        }
    }

    internal class VideoController : IDisposable
    {
        private readonly MediaElement _videoPlayer;
        private readonly DispatcherTimer _timer;
        private bool _isPlayingFlag = false;

        private readonly Func<bool>? _isPlayingGetter;
        private readonly Func<bool>? _stopAtPointEnabledGetter;
        private readonly Func<System.Collections.Generic.IEnumerable<Models.Bookmark>>? _getEntries;
        private readonly Action<Models.Bookmark?>? _onEntryFound;
        private readonly double _stopThresholdSeconds;

        public event EventHandler<VideoPositionEventArgs>? PositionUpdated;

        public VideoController(MediaElement videoPlayer,
                               Func<bool>? isPlayingGetter = null,
                               Func<bool>? stopAtPointEnabledGetter = null,
                               Func<System.Collections.Generic.IEnumerable<Models.Bookmark>>? getEntries = null,
                               Action<Models.Bookmark?>? onEntryFound = null,
                               double stopThresholdSeconds = 0.1)
        {
            _videoPlayer = videoPlayer ?? throw new ArgumentNullException(nameof(videoPlayer));
            _timer = new DispatcherTimer(DispatcherPriority.Normal)
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _timer.Tick += Timer_Tick;

            _isPlayingGetter = isPlayingGetter;
            _stopAtPointEnabledGetter = stopAtPointEnabledGetter;
            _getEntries = getEntries;
            _onEntryFound = onEntryFound;
            _stopThresholdSeconds = stopThresholdSeconds;
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            try
            {
                if (_videoPlayer.NaturalDuration.HasTimeSpan)
                {
                    var current = _videoPlayer.Position;
                    var total = _videoPlayer.NaturalDuration.TimeSpan;
                    bool isPlaying = _isPlayingGetter?.Invoke() ?? _isPlayingFlag;

                    PositionUpdated?.Invoke(this, new VideoPositionEventArgs(current, total, isPlaying));

                    // StopAtPoint logic
                    try
                    {
                        if (( _stopAtPointEnabledGetter?.Invoke() ?? false) && isPlaying && _getEntries != null && _onEntryFound != null)
                        {
                            var entries = _getEntries();
                            var found = System.Linq.Enumerable.OrderBy(entries.Where(bm => System.Math.Abs((bm.Time - current).TotalSeconds) < _stopThresholdSeconds), bm => System.Math.Abs((bm.Time - current).TotalSeconds)).FirstOrDefault();
                            _onEntryFound(found);
                        }
                    }
                    catch
                    {
                        // swallow
                    }
                }
            }
            catch
            {
                // Swallow; keep timer running
            }
        }

        public void SetPlaying(bool isPlaying) => _isPlayingFlag = isPlaying;

        public void Start() => _timer.Start();
        public void Stop() => _timer.Stop();

        // Playback control
        public void Play()
        {
            try { _videoPlayer.Play(); }
            catch { }
            _isPlayingFlag = true;
        }

        public void Pause()
        {
            try { _videoPlayer.Pause(); }
            catch { }
            _isPlayingFlag = false;
        }

        public void StopPlayback()
        {
            try { _videoPlayer.Stop(); }
            catch { }
            _isPlayingFlag = false;
        }

        public void Seek(TimeSpan pos)
        {
            try { _videoPlayer.Position = pos; }
            catch { }
        }

        public void SetSpeed(double speed)
        {
            try { _videoPlayer.SpeedRatio = speed; }
            catch { }
        }

        public void Dispose()
        {
            _timer.Tick -= Timer_Tick;
            _timer.Stop();
        }
    }
}
