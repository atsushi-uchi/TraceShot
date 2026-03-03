using ClosedXML.Excel;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using ScreenRecorderLib;
using System.Diagnostics;
using System.IO;
using System.Text;
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
    public string RecordingTime { get; private set; } = "00:00:00";

    private Recorder? _recorder;
    public List<string> TraceLogs { get; private set; } = new List<string>();
    public string CurrentVideoName { get; private set; } = "";
    public string CurrentFolder { get; set; }  = "";
    public RecordingEvidence? Evidence { get; set; }
    public string? JsonPath { get; set; }
    private List<Bookmark> _currentBookmarks = [];
    public event EventHandler? OnActualRecordingStarted;

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
    public string? SaveSingleBookmarkImage(Bookmark bm, MediaElement videoPlayer, double scale = 0.5)
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
            // ⭐ ポイント：全体にスケーリングを適用
            // これ以降の DrawRectangle 命令はすべて自動的に scale 倍されます
            drawingContext.PushTransform(new ScaleTransform(scale, scale));

            // 背景の描画 (サイズは元の解像度を指定)
            VisualBrush visualBrush = new VisualBrush(videoPlayer) { Stretch = Stretch.Uniform };
            drawingContext.DrawRectangle(visualBrush, null, new Rect(0, 0, originalWidth, originalHeight));

            // 矩形の合成
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

            // PushTransform を閉じます
            drawingContext.Pop();
        }

        // 3. RenderTargetBitmap のサイズを「計算後のサイズ」にする
        RenderTargetBitmap bmp = new RenderTargetBitmap(renderWidth, renderHeight, 96, 96, PixelFormats.Pbgra32);
        bmp.Render(drawingVisual);

        // --- 以下、保存処理は同じ ---
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

    public async Task ExportToPdfAsync(string htmlPath, string pdfPath)
    {
        // 1. ブラウザをダウンロード/管理する準備
        // 最新版では using を使いません
        var browserFetcher = new BrowserFetcher();

        // ブラウザの実行ファイルをダウンロード（初回のみ）
        await browserFetcher.DownloadAsync();

        // 2. ヘッドレスブラウザを起動
        // LaunchAsync には明示的に ExecutablePath を渡すのが現在の推奨です
        using (var browser = await Puppeteer.LaunchAsync(new LaunchOptions
        {
            Headless = true
        }))
        {
            using (var page = await browser.NewPageAsync())
            {
                // 3. HTMLファイルをフルパスで読み込む
                // "file:///" はスラッシュ3つが確実です
                await page.GoToAsync("file:///" + htmlPath.Replace("\\", "/"));

                // 💡 念のため、画像などの読み込み完了を少し待つ
                await page.WaitForNetworkIdleAsync();

                // 4. PDFとして保存
                await page.PdfAsync(pdfPath, new PdfOptions
                {
                    Format = PaperFormat.A4,
                    PrintBackground = true, // 背景色やヘッダーの色を出す
                    MarginOptions = new MarginOptions
                    {
                        Top = "10mm",
                        Bottom = "10mm",
                        Left = "10mm",
                        Right = "10mm"
                    }
                });
            }
        }
    }

    public string ExportToHtml(bool isOriginalSize = false)
    {
        if (Evidence is null) return string.Empty;
        var evidence = Evidence;

        string htmlPath = Path.Combine(CurrentFolder, "EvidenceReport.html");
        StringBuilder sb = new StringBuilder();

        // HTMLのヘッダーとスタイル
        sb.AppendLine("<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>");
        sb.AppendLine("<title>エビデンス報告書</title>");
        // --- スタイル部分の修正 ---
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: sans-serif; margin: 20px; background: #f0f2f5; }");
        sb.AppendLine(".container { max-width: 98%; margin: auto; background: white; padding: 20px; box-shadow: 0 0 15px rgba(0,0,0,0.1); }");
        sb.AppendLine("table { width: 100%; border-collapse: collapse; }");

        // 列幅の設定
        sb.AppendLine(".col-time { width: 80px; text-align: center; }");
        sb.AppendLine(".col-note { width: 150px; }");
        sb.AppendLine(".col-ss { }"); // 画像列

        sb.AppendLine("th, td { border: 1px solid #dee2e6; padding: 10px; vertical-align: top; }");
        sb.AppendLine("th { background-color: #4472C4; color: white; }");

        // 💡 設定によって画像スタイルを切り替える
        if (isOriginalSize)
        {
            // 元のサイズを維持（画面からはみ出る場合はスクロール可能にする）
            sb.AppendLine(".ss-image { display: block; border: 1px solid #ccc; margin: auto; }");
            sb.AppendLine(".col-ss { overflow-x: auto; }");
        }
        else
        {
            // 画面幅にフィットさせる
            sb.AppendLine(".ss-image { width: 100%; height: auto; display: block; border: 1px solid #ccc; }");
        }

        sb.AppendLine("</style>");

        // --- テーブルヘッダー部分の修正 ---
        sb.AppendLine("<table>");
        sb.AppendLine("<tr><th class='col-time'>経過時間</th><th class='col-note'>コメント</th><th class='col-ss'>スクリーンショット</th></tr>");

        // --- データ行部分の修正 ---
        foreach (var bm in evidence.Bookmarks)
        {
            sb.AppendLine("<tr>");
            sb.AppendLine($"<td class='col-time'>{bm.Time}</td>");
            sb.AppendLine($"<td class='col-note'>{bm.Note}</td>");

            string relativePath = $"ScreenShot/{Path.GetFileName(bm.ImagePath)}";
            sb.AppendLine($"<td class='col-ss'><img src='{relativePath}' class='ss-image'></td>");
            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table></div></body></html>");

        File.WriteAllText(htmlPath, sb.ToString());

        // ブラウザで自動表示
        //Process.Start(new ProcessStartInfo(htmlPath) { UseShellExecute = true });
        return htmlPath;
    }

    public void ExportToExcel()
    {
        if (Evidence is null) return;
        var evidence = Evidence;

        string fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + ".xlsx";
        string fullPath = Path.Combine(CurrentFolder, fileName);

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("録画エビデンス");

            // --- 基本設定 ---
            int dataStartRow = 10;
            ws.Cell(dataStartRow, 1).Value = "経過時間";
            ws.Cell(dataStartRow, 2).Value = "内容・メモ";
            ws.Cell(dataStartRow, 3).Value = "スクリーンショット";

            // --- ブックマーク一覧のループ内 ---
            for (int i = 0; i < evidence.Bookmarks.Count; i++)
            {
                var bm = evidence.Bookmarks[i];
                int currentRow = dataStartRow + 1 + i;

                ws.Cell(currentRow, 1).Value = bm.Time;
                ws.Cell(currentRow, 2).Value = bm.Note;

                if (!string.IsNullOrEmpty(bm.ImagePath) && File.Exists(bm.ImagePath))
                {
                    // 1. 画像を追加（この時点ではまだ位置固定しない）
                    var picture = ws.AddPicture(bm.ImagePath);

                    // 2. 💡 MoveTo メソッドを使って、セル位置と「ズレ（オフセット）」を同時に指定する
                    // 第1引数：開始セル
                    // 第2引数：横方向のオフセット（ピクセル）
                    // 第3引数：縦方向のオフセット（ピクセル） ← ここで 10 指定
                    picture.MoveTo(ws.Cell(currentRow, 3), 5, 10);

                    // 3. 等倍に設定
                    picture.Scale(1.0);

                    // 4. 黄金設定の高さ調整
                    double baseHeight = picture.Height * 0.75;
                    double safeBuffer = 80; // さらに余裕を持って 80 に増やしました
                    ws.Row(currentRow).Height = Math.Min(baseHeight + safeBuffer, 409);
                }
            }

            // 💡 仕上げのレイアウト調整
            ws.Columns(1, 2).AdjustToContents(); // テキストに合わせて幅調整
            ws.Column(3).Width = 120; // 画像列は十分な幅を確保
            ws.Rows().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top; // すべて上揃え
            ws.Rows().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // 左揃え

            workbook.SaveAs(fullPath);
        }
    }

    public List<Bookmark> AddBookmark(Bookmark bookmark)
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

    public Bookmark? AddBookmark(string note = " - Screenshot")
    {
        if (_stopwatch.IsRunning)
        {
            var elapsed = _stopwatch.Elapsed;
            string timestamp = elapsed.ToString(@"mm\:ss\.fff");
            var bm = new Bookmark
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
    public string PrepareEvidence(string modeName, string windowTitle)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        string folderName = $"TraceShot_{timestamp}";
        CurrentFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), folderName);

        Directory.CreateDirectory(CurrentFolder);

        CurrentVideoName = $"TraceShot_{timestamp}.mp4";
        string videoPath = Path.Combine(CurrentFolder, CurrentVideoName);

        // JSONもこのタイミングで下書きするか、終了時に保存
        SaveJson(modeName, windowTitle, timestamp);

        return videoPath; // 録画メソッドに渡す用
    }

    public void UpdateJson()
    {
        Evidence.Bookmarks = _currentBookmarks;
        SaveEvidenceJson();
    }

    // 録画完了イベントなどで呼び出す
    private void SaveJson(string mode, string title, string timestamp)
    {
        Evidence = new RecordingEvidence
        {
            VideoFileName = CurrentVideoName,
            RecordingDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
            WindowTitle = title,
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

    public event EventHandler<FrameRecordedEventArgs> OnPreviewFrameReceived;

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
            }
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
            }
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