using ScreenRecorderLib;
using System.Diagnostics;
using System.IO;
using System.Security.RightsManagement;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Brushes = System.Windows.Media.Brushes;
using MessageBox = System.Windows.MessageBox;

public class RecorderManager
{
    private Stopwatch _stopwatch = new Stopwatch();
    private DispatcherTimer _timer;
    private Recorder? _recorder;
    private List<BookMark> _currentBookmarks = [];
    public string RecordingTime { get; private set; } = "00:00:00";
    public List<string> TraceLogs { get; private set; } = new List<string>();
    public string CurrentVideoName { get; private set; } = "";
    public string CurrentFolder { get; set; }  = "";
    public RecordingEvidence? Evidence { get; set; }
    public string? JsonPath { get; set; }
    public int FrameRate { get; set; }
    public bool UseHardwareAccel { get; set; }
    public event EventHandler? OnActualRecordingStarted;
    public event EventHandler<FrameRecordedEventArgs>? OnPreviewFrameReceived;

    public RecorderManager()
    {
        _timer = new DispatcherTimer();
        _timer.Interval = TimeSpan.FromSeconds(1);
        _timer.Tick += (s, e) => {
            RecordingTime = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss");
            Debug.WriteLine($"Recording... {RecordingTime}");
        };
    }

    // 引数に double scale を追加 (例: 0.5 = 50%, 1.0 = 100%)
    public string? SaveSingleBookmarkImage(BookMark bm, MediaElement videoPlayer, double scale = 0.5)
    {
        if (string.IsNullOrEmpty(CurrentFolder)) return null;

        string screenshotFolder = Path.Combine(CurrentFolder, "ScreenShot");
        if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);

        // 1. 動画の本来の解像度を取得
        int originalWidth = videoPlayer.NaturalVideoWidth;
        int originalHeight = videoPlayer.NaturalVideoHeight;
        if (originalWidth == 0 || originalHeight == 0) return null;

        // 2. 出力先の解像度を計算
        int renderWidth = (int)(originalWidth * scale);
        int renderHeight = (int)(originalHeight * scale);

        DrawingVisual drawingVisual = new DrawingVisual();
        using (DrawingContext drawingContext = drawingVisual.RenderOpen())
        {
            // ⭐ ポイント1：全体スケーリングを適用開始
            drawingContext.PushTransform(new ScaleTransform(scale, scale));

            // 배경의描画
            VisualBrush visualBrush = new VisualBrush(videoPlayer) { Stretch = Stretch.Uniform };
            drawingContext.DrawRectangle(visualBrush, null, new Rect(0, 0, originalWidth, originalHeight));

            // 矩形の合成 (ここまでは自動スケーリングでOK)
            if (bm.MarkRects != null && bm.MarkRects.Count > 0)
            {
                foreach (var relRect in bm.MarkRects)
                {
                    Rect scaledRect = new Rect(
                        relRect.X * originalWidth,
                        relRect.Y * originalHeight,
                        relRect.Width * originalWidth,
                        relRect.Height * originalHeight
                    );

                    double penThickness = Math.Max(2.0, originalWidth / 400.0);
                    drawingContext.DrawRectangle(null, new System.Windows.Media.Pen(Brushes.Red, penThickness), scaledRect);
                }
            }

            // ⭐ ポイント2：ここでスケーリングを解除
            drawingContext.Pop();

            // -------------------------------------------------------
            // ⭐ ここから下はスケーリングの『外側』＝出力画像の実際のピクセルサイズで描画
            // -------------------------------------------------------

            // --- バルーンノート (Balloons) の合成 ---
            if (bm.Balloons != null && bm.Balloons.Count > 0)
            {
                foreach (var note in bm.Balloons)
                {
                    // 1. 出力画像（scale適用後）のピクセル座標を計算
                    double outW = originalWidth * scale;
                    double outH = originalHeight * scale;
                    var outputStartPt = new System.Windows.Point(note.TargetPoint.X * outW, note.TargetPoint.Y * outH);
                    var outputEndPt = new System.Windows.Point(note.TextPoint.X * outW, note.TextPoint.Y * outH);

                    // 💡 2. 文字サイズを出力画像の高さ(outH)に合わせて計算 (例: 高さの 3%)
                    // これにより、解像度が変わっても常に「読みやすい大きさ」に固定されます
                    double dynamicFontSize = Math.Max(16.0, outH * 0.03);
                    double thickness = Math.Max(2.0, outW / 500.0);

                    // 3. 線と丸の描画
                    var linePen = new System.Windows.Media.Pen(Brushes.Red, thickness);
                    linePen.DashStyle = new DashStyle(new double[] { 4, 2 }, 0);
                    drawingContext.DrawLine(linePen, outputStartPt, outputEndPt);
                    drawingContext.DrawEllipse(Brushes.Red, null, outputStartPt, thickness * 2, thickness * 2);

                    // 4. FormattedText の作成
                    FormattedText ft = new FormattedText(
                        note.Text,
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Verdana"),
                        dynamicFontSize, // 💡 計算したフォントサイズを適用
                        Brushes.White,
                        VisualTreeHelper.GetDpi(videoPlayer).PixelsPerDip);

                    // 5. 背景とテキストの描画
                    // パディングを少し多めに取る (fontSizeに合わせて調整)
                    double padding = dynamicFontSize * 0.3;
                    Rect textRect = new Rect(outputEndPt.X, outputEndPt.Y, ft.Width + (padding * 2), ft.Height + (padding * 2));

                    drawingContext.DrawRoundedRectangle(
                        new SolidColorBrush(System.Windows.Media.Color.FromArgb(220, 255, 0, 0)), // 少し濃くして視認性アップ
                        null,
                        textRect,
                        padding * 0.5, padding * 0.5);

                    drawingContext.DrawText(ft, new System.Windows.Point(textRect.X + padding, textRect.Y + padding));
                }
            }
        }

        // 3. RenderTargetBitmap のサイズを「計算後のサイズ」にする
        RenderTargetBitmap bmp = new RenderTargetBitmap(renderWidth, renderHeight, 96, 96, PixelFormats.Pbgra32);
        bmp.Render(drawingVisual);

        // --- 保存処理 ---
        string timeStr = bm.Time.Replace(":", "");
        string fileName = $"SS_{timeStr}.png";
        string filePath = Path.Combine(screenshotFolder, fileName);

        using (FileStream fs = new FileStream(filePath, FileMode.Create))
        {
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bmp));
            encoder.Save(fs);
        }

        return filePath;
    }


    public List<BookMark> AddBookmark(BookMark bookmark)
    {
        _currentBookmarks.Add(bookmark);
        var sorted = _currentBookmarks.OrderBy(b => b.Seconds).ToList();
        _currentBookmarks.Clear();
        foreach (var b in sorted)
        {
            _currentBookmarks.Add(b);
        }
        if (Evidence is not null) Evidence.Bookmarks = _currentBookmarks;
        return _currentBookmarks;
    }

    public BookMark? AddBookmark(string note = " - Screenshot")
    {
        if (_stopwatch.IsRunning)
        {
            var elapsed = _stopwatch.Elapsed;
            string timestamp = elapsed.ToString(@"mm\:ss\.fff");
            var bm = new BookMark
            {
                Time = timestamp,
                Seconds = elapsed.TotalSeconds,
                Note = note,
            };
            _currentBookmarks.Add(bm);

            if (Evidence is not null) Evidence.Bookmarks = _currentBookmarks;

            return bm;
        }
        return null;
    }

    public string PrepareEvidence(string path, string modeName)
    {
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
        string folderName = $"TraceShot_{timestamp}";
        CurrentFolder = Path.Combine(path, folderName);

        Directory.CreateDirectory(CurrentFolder);

        CurrentVideoName = $"TraceShot_{timestamp}.mp4";
        string videoPath = Path.Combine(CurrentFolder, CurrentVideoName);

        SaveJson(modeName, timestamp);

        return videoPath;
    }

    public void UpdateJson()
    {
        Evidence.Bookmarks = _currentBookmarks;
        SaveEvidenceJson();
    }

    // 録画完了イベントなどで呼び出す
    private void SaveJson(string mode, string timestamp)
    {
        Evidence = new RecordingEvidence
        {
            VideoFileName = CurrentVideoName,
            RecordingDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
            Mode = mode,
            Bookmarks = _currentBookmarks,
        };

        JsonPath = Path.Combine(CurrentFolder, $"TraceShot_{timestamp}.json");
        SaveEvidenceJson();
    }

    public void SaveEvidenceJson()
    {
        if (JsonPath == null)
        {
            MessageBox.Show("保存データがありません");
            return;
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
        };
        string jsonString = JsonSerializer.Serialize(Evidence, options);

        File.WriteAllText(JsonPath, jsonString);
    }

    public void StartFullscreenRecording(string filePath, string targetDeviceName)
    {
        TraceLogs.Clear();
        _currentBookmarks.Clear();

        DisplayRecordingSource? screenSource = null;
        if (string.IsNullOrEmpty(targetDeviceName))
        {
            var displays = Recorder.GetDisplays();
            screenSource = displays.FirstOrDefault();
        }
        else
        {
            screenSource = new DisplayRecordingSource(targetDeviceName);
        }

        if (screenSource == null)
        {
            MessageBox.Show("モニターを選択してください。");
            return;
        }

        var options = new RecorderOptions
        {
            SourceOptions = new SourceOptions
            {
                RecordingSources = { screenSource }
            },
            OutputOptions = new OutputOptions
            {
                RecorderMode = RecorderMode.Video,
                IsVideoFramePreviewEnabled = true,
            },
            VideoEncoderOptions = new VideoEncoderOptions
            {
                Framerate = FrameRate,
                IsHardwareEncodingEnabled = UseHardwareAccel,
            },
        };

        // 3. インスタンス生成と開始
        _recorder = Recorder.CreateRecorder(options);
        _recorder.OnFrameRecorded += (s, e) => OnPreviewFrameReceived?.Invoke(this, e);

        // ステータス変更イベントを登録
        _recorder.OnStatusChanged += (s, e) =>
        {
            // ステータスが "Recording" になったら準備完了
            if (_recorder.Status == RecorderStatus.Recording)
            {
                // UIスレッドなどでタイマーを開始させるための通知を送る
                OnActualRecordingStarted?.Invoke(this, EventArgs.Empty);
            }
        };
        _recorder.OnRecordingComplete += (s, e) => TraceLogs.Add("Recording Complete");
        _recorder.OnRecordingFailed += (s, e) => TraceLogs.Add("Recording Failed: " + e.Error);

        _timer.Start();
        _stopwatch.Restart();
        _recorder.Record(filePath);
    }

    public void StartRectangleRecording(string filePath, string targetDeviceName, Rectangle? region = null)
    {
        TraceLogs.Clear();
        _currentBookmarks.Clear();

        // 1. ソースの作成
        DisplayRecordingSource? screenSource = null;
        if (string.IsNullOrEmpty(targetDeviceName))
        {
            var displays = Recorder.GetDisplays();
            screenSource = displays.FirstOrDefault();
        }
        else
        {
            screenSource = new DisplayRecordingSource(targetDeviceName);
        }

        // 矩形選択時の座標設定
        screenSource.SourceRect = new ScreenRect(
            region.Value.Left, region.Value.Top, region.Value.Width, region.Value.Height);

        // 2. オプションの組み立て
        var options = new RecorderOptions
        {
            SourceOptions = new SourceOptions
            {
                // RecordingSources プロパティを使用（中括弧 { } で追加）
                RecordingSources = { screenSource }
            },
            OutputOptions = new OutputOptions
            {
                RecorderMode = RecorderMode.Video,
                IsVideoFramePreviewEnabled = true,
                OutputFrameSize = new ScreenSize(region.Value.Width, region.Value.Height),
            },
            VideoEncoderOptions = new VideoEncoderOptions
            {
                Framerate = FrameRate,
                IsHardwareEncodingEnabled = UseHardwareAccel,
            },
        };

        // 3. インスタンス生成と開始
        _recorder = Recorder.CreateRecorder(options);
        _recorder.OnFrameRecorded += (s, e) => OnPreviewFrameReceived?.Invoke(this, e);

        // ステータス変更イベントを登録
        _recorder.OnStatusChanged += (s, e) =>
        {
            // ステータスが "Recording" になったら準備完了
            if (_recorder.Status == RecorderStatus.Recording)
            {
                // UIスレッドなどでタイマーを開始させるための通知を送る
                OnActualRecordingStarted?.Invoke(this, EventArgs.Empty);
            }
        };
        _recorder.OnRecordingComplete += (s, e) => TraceLogs.Add("Recording Complete");
        _recorder.OnRecordingFailed += (s, e) => TraceLogs.Add("Recording Failed: " + e.Error);

        _timer.Start();
        _stopwatch.Restart();
        _recorder.Record(filePath);
    }

    public void StartWindowRecording(string filePath, IntPtr windowHandle)
    {
        TraceLogs.Clear();

        // 1. ウィンドウをソースとして作成
        var windowSource = new WindowRecordingSource(windowHandle);

        // 2. オプションの組み立て
        var options = new RecorderOptions
        {
            SourceOptions = new SourceOptions
            {
                // コレクション初期化子で追加
                RecordingSources = { windowSource }
            },
            OutputOptions = new OutputOptions
            {
                RecorderMode = RecorderMode.Video,
                IsVideoFramePreviewEnabled = true,
                OutputFrameSize = null,
            },
            VideoEncoderOptions = new VideoEncoderOptions
            {
                Framerate = FrameRate,
                IsHardwareEncodingEnabled = UseHardwareAccel,
            },
        };

        // 3. インスタンス生成と開始
        _recorder = Recorder.CreateRecorder(options);
        _recorder.OnFrameRecorded += (s, e) => OnPreviewFrameReceived?.Invoke(this, e);
        _recorder.OnRecordingComplete += (s, e) => TraceLogs.Add("Window Recording Complete");
        _recorder.OnRecordingFailed += (s, e) => TraceLogs.Add("Window Recording Failed: " + e.Error);

        _timer.Start();
        _stopwatch.Restart();
        _recorder.Record(filePath);
    }

    public void StopRecording()
    {
        _timer.Stop();
        _stopwatch.Stop();
        _recorder?.Stop();
        SaveEvidenceJson();
    }
}