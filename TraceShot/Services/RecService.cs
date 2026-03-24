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
using TraceShot.Controls;
using TraceShot.Models;
using Brushes = System.Windows.Media.Brushes;
using Pen = System.Windows.Media.Pen;
using Point = System.Windows.Point;

namespace TraceShot.Services
{
    public partial class RecService : ObservableObject
    {
        public static RecService Instance { get; } = new RecService();

        [ObservableProperty][NotifyPropertyChangedFor(nameof(Entries))]
        private RecEvidence _evidence = new();

        public ObservableCollection<Bookmark> Entries => Evidence.Entries;

        [ObservableProperty]
        private bool _isRecording = false;

        [ObservableProperty]
        private string _recordingTime = "00:00:00";

        private Stopwatch _stopwatch = new Stopwatch();
        private DispatcherTimer _timer;
        private Recorder? _recorder;
        private DateTime _actualStartTime;
        public List<string> TraceLogs { get; private set; } = [];
        public TimeSpan CurrentDuration => _stopwatch.Elapsed;
        public string CurrentVideoName { get; private set; } = "";
        public string CurrentFolder { get; set; } = "";
        public string? JsonPath { get; set; }
        public int FrameRate { get; set; }
        public bool UseHardwareAccel { get; set; }

        public event EventHandler? OnActualRecordingStarted;
        public event EventHandler<FrameRecordedEventArgs>? OnPreviewFrameReceived;

        // 録画停止時に実行する処理を登録するためのアクション
        public Action? OnRecordingStopped { get; set; }

        private RecService()
        {
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(0.5);
            _timer.Tick += (s, e) =>
            {
                RecordingTime = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss");
            };
        }

        public Bookmark? GetBookmark(TimeSpan time)
        {

            return Entries.FirstOrDefault(b => Math.Abs(b.Time.TotalSeconds - time.TotalSeconds) < 0.1);
        }

        public Bookmark? GetBookmark(AnnotationBase annotation)
        {
            return Entries.Where(e => e.Annotations.Contains(annotation)).FirstOrDefault();
        }

        public string? SaveBitmap(Bookmark bm, WriteableBitmap source)
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
                return filePath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Backup Save Error: {ex.Message}");
                return null;
            }
        }


        public (string? Path, BitmapSource? Bitmap)? SaveImage(Bookmark bm, VideoSnapshotInfo? info, double scale = 0.5, bool saveToFile = true)
        {
            if (string.IsNullOrEmpty(CurrentFolder) || info is null) return null;

            string screenshotFolder = Path.Combine(CurrentFolder, "ScreenShot");
            if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);

            // 1. 基本情報の取得
            int originalWidth = info.NaturalWidth;
            int originalHeight = info.NaturalHeight;
            if (originalWidth == 0 || originalHeight == 0) return (null, null);

            // 個別フォーカスがあれば使用、クロップ有効なら設定値を使用、無効なら全範囲(0,0,1,1)
            Rect cropRel;
            var focus = bm.Annotations.Where(x => x is RectAnnotation rect && rect.IsFocused).FirstOrDefault();
            if (focus != null)
            {
                cropRel = new Rect(focus.RelX, focus.RelY, focus.RelWidth, focus.RelHeight);
            }
            else if (Evidence.IsCropEnabled)
            {
                cropRel = Evidence.CommonCropRect;
            }
            else
            {
                cropRel = new Rect(0, 0, 1, 1);
            }

            // 3. 最終的な出力ピクセルサイズを計算 (クロップ範囲 × スケール)
            int renderWidth = (int)(originalWidth * cropRel.Width * scale);
            int renderHeight = (int)(originalHeight * cropRel.Height * scale);

            if (renderWidth <= 0 || renderHeight <= 0) return (null, null);

            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                // --- A. 動画・矩形用の座標変換 (Scale 内側) ---
                var group = new TransformGroup();
                // クロップの左上(Pixel)を (0,0) に持ってくるためにマイナス移動
                group.Children.Add(new TranslateTransform(-originalWidth * cropRel.X, -originalHeight * cropRel.Y));
                // スケーリングを適用
                group.Children.Add(new ScaleTransform(scale, scale));

                drawingContext.PushTransform(group);

                // 動画の描画
                drawingContext.DrawRectangle(info.VideoBrush, null, new Rect(0, 0, originalWidth, originalHeight));

                // 矩形注釈の描画
                var rects = bm.Rects.OfType<RectAnnotation>();
                foreach (var rectAnno in rects)
                {
                    if (rectAnno.IsFocused) continue;

                    Rect scaledRect = new Rect(
                        rectAnno.RelX * originalWidth,
                        rectAnno.RelY * originalHeight,
                        rectAnno.RelWidth * originalWidth,
                        rectAnno.RelHeight * originalHeight
                    );
                    double penThickness = Math.Max(2.0, originalWidth / 400.0);

                    if (rectAnno.IsMasking)
                    {
                        drawingContext.DrawRectangle(Brushes.Black, null, scaledRect);
                    }
                    else
                    {
                        Pen pen = new Pen(SettingsService.Instance.MainBrush, penThickness);
                        drawingContext.DrawRectangle(null, pen, scaledRect);
                    }
                }

                drawingContext.Pop(); // Transform解除

                //---B.バルーンノートの合成(Scale 外側 / 出力サイズ基準)-- -
                foreach (var noteAnno in bm.Notes)
                {
                    // 全体に対する相対比率 (0.0~1.0)
                    double ratioX = noteAnno.X / info.ActualViewWidth;
                    double ratioY = noteAnno.Y / info.ActualViewHeight;
                    double ratioStartX = noteAnno.StartX / info.ActualViewWidth;
                    double ratioStartY = noteAnno.StartY / info.ActualViewHeight;

                    // クロップ範囲内での位置に変換し、出力サイズを掛ける
                    // (全体比率 - クロップ開始位置) / クロップ幅 = クロップ内での相対位置
                    var outputEndPt = new Point(
                        (ratioX - cropRel.X) / cropRel.Width * renderWidth,
                        (ratioY - cropRel.Y) / cropRel.Height * renderHeight
                    );
                    var outputStartPt = new Point(
                        (ratioStartX - cropRel.X) / cropRel.Width * renderWidth,
                        (ratioStartY - cropRel.Y) / cropRel.Height * renderHeight
                    );

                    // フォント・サイズの計算
                    double screenToOutputRatio = (originalWidth * scale) / info.ActualViewWidth;
                    double dynamicFontSize = Math.Max(16.0, (originalHeight * scale * cropRel.Height) * 0.03);
                    double padding = dynamicFontSize * 0.5;
                    double thickness = Math.Max(2.0, renderWidth / 500.0);

                    FormattedText ft = new FormattedText(
                        noteAnno.Text ?? "",
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Verdana"),
                        dynamicFontSize,
                        SettingsService.Instance.MainTextBrush,
                        VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                    ft.MaxTextWidth = Math.Max(100, 200 * screenToOutputRatio);

                    double boxWidth = ft.Width + (padding * 2);
                    double boxHeight = ft.Height + (padding * 2);

                    Rect textRect = new Rect(
                        outputEndPt.X - (boxWidth / 2.0),
                        outputEndPt.Y - (boxHeight / 2.0),
                        boxWidth, boxHeight);

                    // 下地（線・丸・背景箱）
                    var linePen = new Pen(SettingsService.Instance.MainBrush, thickness) { DashStyle = new DashStyle(new double[] { 4, 2 }, 0) };
                    drawingContext.DrawLine(linePen, outputStartPt, outputEndPt);
                    drawingContext.DrawEllipse(SettingsService.Instance.MainBrush, null, outputStartPt, thickness * 2, thickness * 2);

                    drawingContext.DrawRoundedRectangle(
                        SettingsService.Instance.MainBrush,
                        null, textRect, padding * 0.5, padding * 0.5);

                    // 袋文字
                    var textPos = new Point(textRect.X + padding, textRect.Y + padding);
                    Geometry textGeometry = ft.BuildGeometry(textPos);
                    var outlinePen = new Pen(Brushes.Black, dynamicFontSize * 0.15) { LineJoin = PenLineJoin.Round };

                    drawingContext.DrawGeometry(null, outlinePen, textGeometry);
                    drawingContext.DrawGeometry(Brushes.White, null, textGeometry);
                }
            }

            // 4. クロップ後のサイズでビットマップ化
            RenderTargetBitmap bmp = new RenderTargetBitmap(renderWidth, renderHeight, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(drawingVisual);

            // --- 保存処理 ---
            DateTime startDate = Evidence?.RecordingDate ?? DateTime.Now;
            DateTime timestamp = startDate.Add(bm.Time);
            string fileName = $"SS_{timestamp:yyyy-MM-dd_HHmmss_fff}.png";
            string filePath = Path.Combine(screenshotFolder, fileName);

            if (saveToFile)
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Create))
                {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bmp));
                    encoder.Save(fs);
                }
            }
            return (filePath, bmp);
        }

        public void AddBookmark(Bookmark bookmark)
        {
            if (Evidence == null) return;

            Entries.Add(bookmark);
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
                Entries.Add(bm);
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

            Entries.Clear();
            TraceLogs.Clear();

            _recorder.OnRecordingComplete += (s, e) => TraceLogs.Add("Window Recording Complete");
            _recorder.OnRecordingFailed += (s, e) =>
            {
                var errorMsg = "Recording Failed: " + e.Error;
                TraceLogs.Add(errorMsg);
                Debug.WriteLine(errorMsg); // ★ 出力ウィンドウに表示
                System.Windows.MessageBox.Show(errorMsg); // ★ 実行中に確実に気づけるようにする
            };

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

            RecordingTime = "00:00:00";

            // 登録されている処理（MainWindowのモード切替など）を実行
            OnRecordingStopped?.Invoke();
        }
    }
}