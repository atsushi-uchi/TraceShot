using CommunityToolkit.Mvvm.ComponentModel;
using ScreenRecorderLib;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using TraceShot.Models;
using Brushes = System.Windows.Media.Brushes;

namespace TraceShot.Services
{
    public partial class RecManager : ObservableObject
    {
        public static RecManager Instance { get; } = new RecManager();

        [ObservableProperty][NotifyPropertyChangedFor(nameof(Bookmarks))]
        private RecEvidence _evidence = new();

        public ObservableCollection<Bookmark> Bookmarks => Evidence.Bookmarks;

        private Stopwatch _stopwatch = new Stopwatch();
        private DispatcherTimer _timer;
        private Recorder? _recorder;
        private DateTime _actualStartTime;
        [ObservableProperty]
        private bool _isRecording = true;
        [ObservableProperty]
        private string _recordingTime = "00:00:00";
        public List<string> TraceLogs { get; private set; } = [];
        public TimeSpan CurrentDuration => _stopwatch.Elapsed;
        public string CurrentVideoName { get; private set; } = "";
        public string CurrentFolder { get; set; } = "";
        public string? JsonPath { get; set; }
        public int FrameRate { get; set; }
        public bool UseHardwareAccel { get; set; }
        public bool IsCropLocked { get; set; }
        public event EventHandler? OnActualRecordingStarted;
        public event EventHandler<FrameRecordedEventArgs>? OnPreviewFrameReceived;

        private RecManager()
        {
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(0.5);
            _timer.Tick += (s, e) =>
            {
                RecordingTime = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss");
            };
        }

        public string? SaveBackupFromWriteableBitmap(Bookmark bm, WriteableBitmap source)
        {
            if (string.IsNullOrEmpty(CurrentFolder) || source == null) return null;

            string screenshotFolder = Path.Combine(CurrentFolder, "ScreenShot");
            if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);

            var timestamp = _actualStartTime.Add(bm.Time);
            string fileName = $"SS_{timestamp:yyyy-MM-dd_HHmmss_fff}.png";
            string filePath = Path.Combine(screenshotFolder, fileName);

            try
            {
                // 💡 2. UIスレッドのデータを安全にコピー
                var bitmapCopy = source.Clone();
                bitmapCopy.Freeze();

                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapCopy));
                    encoder.Save(stream);
                    stream.Flush();
                }

                Debug.WriteLine($"Backup success: {filePath}");
                return filePath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Backup Save Error: {ex.Message}");
                return null;
            }
        }

        // 引数に double scale を追加 (例: 0.5 = 50%, 1.0 = 100%)
        public (string? Path, BitmapSource? Bitmap)? SaveSingleBookmarkImage(Bookmark bm, VideoSnapshotInfo? info, double scale = 0.5)
        {
            if (string.IsNullOrEmpty(CurrentFolder)) return (null, null);

            string screenshotFolder = Path.Combine(CurrentFolder, "ScreenShot");
            if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);

            // 1. infoクラスから解像度を取得
            int originalWidth = info.NaturalWidth;
            int originalHeight = info.NaturalHeight;
            if (originalWidth == 0 || originalHeight == 0) return (null, null);

            int renderWidth = (int)(originalWidth * scale);
            int renderHeight = (int)(originalHeight * scale);

            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                // ⭐ ポイント1：全体スケーリングを適用開始
                drawingContext.PushTransform(new ScaleTransform(scale, scale));

                // 2. infoクラスの VisualBrush を使用
                drawingContext.DrawRectangle(info.VideoBrush, null, new Rect(0, 0, originalWidth, originalHeight));

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
                foreach (var note in bm.Balloons)
                {
                    // --- 1. すべての計算を最初に行う ---
                    double outW = originalWidth * scale;
                    double outH = originalHeight * scale;
                    var outputStartPt = new System.Windows.Point(note.TargetPoint.X * outW, note.TargetPoint.Y * outH);
                    var outputEndPt = new System.Windows.Point(note.TextPoint.X * outW, note.TextPoint.Y * outH);

                    double dynamicFontSize = Math.Max(16.0, outH * 0.03);
                    double padding = dynamicFontSize * 0.3;
                    double thickness = Math.Max(2.0, outW / 500.0);

                    // FormattedText を作ってサイズを確定させる
                    FormattedText ft = new FormattedText(
                        note.Text,
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Verdana"),
                        dynamicFontSize,
                        Brushes.White,
                        info.DpiScale);

                    // 背景矩形のサイズを確定させる
                    Rect textRect = new Rect(outputEndPt.X, outputEndPt.Y, ft.Width + (padding * 2), ft.Height + (padding * 2));

                    // --- 2. 下地（線・丸・背景箱）を描画する ---
                    // 線と丸
                    var linePen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Red, thickness);
                    linePen.DashStyle = new System.Windows.Media.DashStyle(new double[] { 4, 2 }, 0);
                    drawingContext.DrawLine(linePen, outputStartPt, outputEndPt);
                    drawingContext.DrawEllipse(System.Windows.Media.Brushes.Red, null, outputStartPt, thickness * 2, thickness * 2);

                    // 赤い背景箱
                    drawingContext.DrawRoundedRectangle(
                        new SolidColorBrush(System.Windows.Media.Color.FromArgb(220, 255, 0, 0)),
                        null, textRect, padding * 0.5, padding * 0.5);

                    // --- 3. 最前面に袋文字を描画する ---
                    var textPos = new System.Windows.Point(textRect.X + padding, textRect.Y + padding);
                    Geometry textGeometry = ft.BuildGeometry(textPos);

                    // 黒い縁取り
                    double outlineThickness = dynamicFontSize * 0.15;
                    var outlinePen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Black, outlineThickness)
                    {
                        LineJoin = PenLineJoin.Round
                    };
                    drawingContext.DrawGeometry(null, outlinePen, textGeometry);

                    // 白い中身
                    drawingContext.DrawGeometry(System.Windows.Media.Brushes.White, null, textGeometry);

                    // ※
                }
            }

            // 3. RenderTargetBitmap のサイズを「計算後のサイズ」にする
            RenderTargetBitmap bmp = new RenderTargetBitmap(renderWidth, renderHeight, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(drawingVisual);

            // --- 保存処理 ---
            DateTime startDate = Evidence?.RecordingDate ?? DateTime.Now;
            DateTime timestamp = startDate.Add(bm.Time);

            string fileName = $"SS_{timestamp:yyyy-MM-dd_HHmmss_fff}.png";
            string filePath = Path.Combine(screenshotFolder, fileName);

            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bmp));
                encoder.Save(fs);
                fs.Flush();
            }
            return (filePath, bmp);
        }

        public (string? Path, BitmapSource? Bitmap)? SaveCroppedBookmarkImage(Bookmark bm, VideoSnapshotInfo info, MarkRect cropRectRel, double scale = 1.0)
        {
            if (string.IsNullOrEmpty(CurrentFolder)) return null;

            string screenshotFolder = Path.Combine(CurrentFolder, "ScreenShot");
            if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);

            // 1. info クラスから元の解像度を取得
            int originalWidth = info.NaturalWidth;
            int originalHeight = info.NaturalHeight;
            if (originalWidth == 0 || originalHeight == 0) return null;

            // クロップ範囲のピクセルサイズを計算
            Rect cropRectPix = new Rect(
                cropRectRel.X * originalWidth,
                cropRectRel.Y * originalHeight,
                cropRectRel.Width * originalWidth,
                cropRectRel.Height * originalHeight
            );

            // 出力先のサイズ（クロップ範囲 × スケール）
            int renderWidth = (int)(cropRectPix.Width * scale);
            int renderHeight = (int)(cropRectPix.Height * scale);
            if (renderWidth <= 0 || renderHeight <= 0) return null;

            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                // --- 描画キャンバスのトランスフォーム設定 ---
                // 1. 全体スケールを適用
                drawingContext.PushTransform(new ScaleTransform(scale, scale));
                // 2. 平移: クロップ範囲の左上(X, Y)が(0, 0)に来るように、全体を負の方向にずらす
                drawingContext.PushTransform(new TranslateTransform(-cropRectPix.X, -cropRectPix.Y));

                // 動画全体の描画（info.VideoBrush を使用）
                drawingContext.DrawRectangle(info.VideoBrush, null, new Rect(0, 0, originalWidth, originalHeight));

                // 矩形（赤枠）の合成
                if (bm.MarkRects != null)
                {
                    foreach (var relRect in bm.MarkRects)
                    {
                        if (relRect.IsCropArea) continue;
                        Rect scaledRect = new Rect(
                            relRect.X * originalWidth, relRect.Y * originalHeight,
                            relRect.Width * originalWidth, relRect.Height * originalHeight
                        );
                        double penThickness = Math.Max(2.0, originalWidth / 400.0);
                        drawingContext.DrawRectangle(null, new System.Windows.Media.Pen(System.Windows.Media.Brushes.Red, penThickness), scaledRect);
                    }
                }

                drawingContext.Pop(); // TranslateTransform を解除
                drawingContext.Pop(); // ScaleTransform を解除

                // --- バルーンノートの描画 (出力ピクセル基準) ---
                foreach (var note in bm.Balloons)
                {
                    // クロップ後の座標系における相対位置を計算
                    var croppedRelTargetPt = new System.Windows.Point(
                        (note.TargetPoint.X * originalWidth - cropRectPix.X) / cropRectPix.Width,
                        (note.TargetPoint.Y * originalHeight - cropRectPix.Y) / cropRectPix.Height
                    );
                    var croppedRelTextPt = new System.Windows.Point(
                        (note.TextPoint.X * originalWidth - cropRectPix.X) / cropRectPix.Width,
                        (note.TextPoint.Y * originalHeight - cropRectPix.Y) / cropRectPix.Height
                    );

                    double outW = renderWidth;
                    double outH = renderHeight;
                    var outputStartPt = new System.Windows.Point(croppedRelTargetPt.X * outW, croppedRelTargetPt.Y * outH);
                    var outputEndPt = new System.Windows.Point(croppedRelTextPt.X * outW, croppedRelTextPt.Y * outH);

                    double dynamicFontSize = Math.Max(16.0, outH * 0.03);
                    double padding = dynamicFontSize * 0.3;
                    double thickness = Math.Max(2.0, outW / 500.0);

                    FormattedText ft = new FormattedText(
                        note.Text,
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Verdana"),
                        dynamicFontSize,
                        System.Windows.Media.Brushes.White,
                        info.DpiScale); // info から取得

                    Rect textRect = new Rect(outputEndPt.X, outputEndPt.Y, ft.Width + (padding * 2), ft.Height + (padding * 2));

                    // 背景とラインの描画
                    var linePen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Red, thickness);
                    linePen.DashStyle = new System.Windows.Media.DashStyle(new double[] { 4, 2 }, 0);
                    drawingContext.DrawLine(linePen, outputStartPt, outputEndPt);
                    drawingContext.DrawEllipse(System.Windows.Media.Brushes.Red, null, outputStartPt, thickness * 2, thickness * 2);

                    drawingContext.DrawRoundedRectangle(new SolidColorBrush(System.Windows.Media.Color.FromArgb(220, 255, 0, 0)), null, textRect, padding * 0.5, padding * 0.5);

                    // 袋文字の描画
                    var textPos = new System.Windows.Point(textRect.X + padding, textRect.Y + padding);
                    Geometry textGeometry = ft.BuildGeometry(textPos);
                    drawingContext.DrawGeometry(null, new System.Windows.Media.Pen(System.Windows.Media.Brushes.Black, dynamicFontSize * 0.15) { LineJoin = PenLineJoin.Round }, textGeometry);
                    drawingContext.DrawGeometry(System.Windows.Media.Brushes.White, null, textGeometry);
                }
            }

            // 画像生成
            RenderTargetBitmap bmp = new RenderTargetBitmap(renderWidth, renderHeight, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(drawingVisual);

            // ファイル保存
            DateTime startDate = Evidence?.RecordingDate ?? DateTime.Now;
            DateTime timestamp = startDate.Add(bm.Time);
            string fileName = $"SS_Crop_{timestamp:yyyy-MM-dd_HHmmss_fff}.png";
            string filePath = Path.Combine(screenshotFolder, fileName);

            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bmp));
                encoder.Save(fs);
            }

            return (filePath, bmp);
        }

        public void AddBookmark(Bookmark bookmark)
        {
            if (Evidence == null) return;

            Bookmarks.Add(bookmark);
        }

        public Bookmark? AddBookmark(string note = " - Screenshot")
        {
            if (Evidence == null) return null;

            if (_stopwatch.IsRunning)
            {
                var bm = new Bookmark
                {
                    Time = _stopwatch.Elapsed,
                    Icon = "📌",
                    Note = note,
                };
                Bookmarks.Add(bm);
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


        // 録画完了イベントなどで呼び出す
        private void SaveJson(string mode, string timestamp)
        {
            Evidence.VideoFileName = CurrentVideoName;
            Evidence.RecordingDate = DateTime.Now;
            Evidence.RecMode = mode;
            Evidence.IsCropLocked = IsCropLocked;

            JsonPath = Path.Combine(CurrentFolder, $"TraceShot_{timestamp}.json");
            SaveEvidenceJson();
        }

        public void SaveEvidenceJson()
        {
            if (JsonPath == null) return;

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
                throw new Exception("録画するモニターが見つかりません。");
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

            StartRecording(filePath);
        }

        public void StartRectangleRecording(string filePath, string targetDeviceName, Rectangle? region)
        {
            if (region is null) return;

            TraceLogs.Clear();

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
            if (screenSource != null)
            {
                screenSource!.SourceRect = new ScreenRect(
                    region.Value.Left, region.Value.Top, region.Value.Width, region.Value.Height);
            }

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

            StartRecording(filePath);
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

            StartRecording(filePath);
        }

        private void StartRecording(string filePath)
        {
            if (_recorder is null) return;
            IsRecording = true;

            Bookmarks.Clear();
            TraceLogs.Clear();

            _recorder.OnRecordingComplete += (s, e) => TraceLogs.Add("Window Recording Complete");
            _recorder.OnRecordingFailed += (s, e) => TraceLogs.Add("Window Recording Failed: " + e.Error);

            _actualStartTime = Evidence?.RecordingDate ?? DateTime.Now;

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
            IsRecording = false;
        }
    }
}