using ClosedXML.Excel;
using Microsoft.Win32;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;

namespace TraceShot.Features
{
    /// <summary>
    /// ExportWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ExportWindow : Window
    {
        public ExportWindow()
        {
            InitializeComponent();
        }

        private double GetSelectedScale()
        {
            if (ExportScaleComboBox.SelectedItem is ComboBoxItem item &&
                double.TryParse(item.Tag.ToString(), out double scale))
            {
                return scale;
            }
            return 1.0; // デフォルト
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            // フォルダ選択ダイアログのインスタンスを作成
            var dialog = new OpenFolderDialog
            {
                Title = "出力先フォルダを選択してください",
                InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
            };

            // ダイアログを表示
            if (dialog.ShowDialog() == true)
            {
                // 選択されたパスを TextBox にセット
                OutputPathBox.Text = dialog.FolderName;
            }
        }

        private async void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var main = Owner as MainWindow;
            if (main?.RecorderMgr.Evidence == null)
            {
                MessageBox.Show("エクスポートするデータがありません。");
                return;
            }

            try
            {
                foreach (var bm in main.RecorderMgr.Evidence.Bookmarks)
                {
                    main.VideoPlayer.Position = TimeSpan.FromSeconds(bm.Seconds);
                    await Task.Delay(500);

                    // 画像を保存し、そのパスを bm.ImagePath に格納するようマネージャー側を調整
                    var scale = GetSelectedScale();
                    var savedPath = main.RecorderMgr.SaveSingleBookmarkImage(bm, main.VideoPlayer, scale);
                    bm.ImagePath = savedPath;
                }
                ExportToExcel();
                MessageBox.Show("Excelを出力しました！");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エクスポート失敗: {ex.Message}");
            }
        }

        private async void ExportHtml_Click(object sender, RoutedEventArgs e)
        {
            var main = Owner as MainWindow;
            if (main?.RecorderMgr.Evidence == null)
            {
                MessageBox.Show("エクスポートするデータがありません。");
                return;
            }
            var evidence = main.RecorderMgr.Evidence;

            try
            {
                foreach (var bm in evidence.Bookmarks)
                {
                    main.VideoPlayer.Position = TimeSpan.FromSeconds(bm.Seconds);
                    await Task.Delay(500);
                    var scale = GetSelectedScale();
                    var savedPath = main.RecorderMgr.SaveSingleBookmarkImage(bm, main.VideoPlayer, scale);
                    bm.ImagePath = savedPath;
                }

                var sb = new StringBuilder();

                // --- HTML ヘッダー・スタイル部分は変更なし ---
                sb.AppendLine("<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>");
                sb.AppendLine("<title>エビデンス報告書</title>");
                sb.AppendLine("<style>");
                sb.AppendLine("body { font-family: sans-serif; margin: 20px; background: #f0f2f5; }");
                sb.AppendLine(".container { max-width: 98%; margin: auto; background: white; padding: 20px; box-shadow: 0 0 15px rgba(0,0,0,0.1); }");
                sb.AppendLine("table { width: 100%; border-collapse: collapse; }");
                sb.AppendLine(".col-time { width: 80px; text-align: center; }");
                sb.AppendLine(".col-note { width: 150px; }");
                sb.AppendLine("th, td { border: 1px solid #dee2e6; padding: 10px; vertical-align: top; }");
                sb.AppendLine("th { background-color: #4472C4; color: white; }");

                sb.AppendLine(".ss-image { width: 100%; height: auto; display: block; border: 1px solid #ccc; }");
                sb.AppendLine("</style></head><body><div class='container'>");

                sb.AppendLine("<table>");
                sb.AppendLine("<tr><th class='col-time'>経過時間</th><th class='col-note'>コメント</th><th class='col-ss'>スクリーンショット</th></tr>");

                // --- データ行部分の修正：ここがポイント ---
                foreach (var bm in evidence.Bookmarks)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td class='col-time'>{bm.Time}</td>");
                    sb.AppendLine($"<td class='col-note'>{bm.Note}</td>");

                    // 画像ファイルを読み込んで Base64 に変換
                    string base64Image = "";
                    if (File.Exists(bm.ImagePath))
                    {
                        byte[] imageBytes = File.ReadAllBytes(bm.ImagePath);
                        string base64String = Convert.ToBase64String(imageBytes);
                        // png か jpg かは拡張子から判断（ここでは汎用的に png 指定でも大抵動きます）
                        base64Image = $"data:image/png;base64,{base64String}";
                    }

                    // src に相対パスではなく Base64 文字列を入れる
                    sb.AppendLine($"<td class='col-ss'><img src='{base64Image}' class='ss-image'></td>");
                    sb.AppendLine("</tr>");
                }

                sb.AppendLine("</table></div></body></html>");

                var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + "_full.html";
                var fullPath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                    main.RecorderMgr.CurrentFolder : OutputPathBox.Text, fileName);
                File.WriteAllText(fullPath, sb.ToString());

                MessageBox.Show("HTML出力");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"出力エラー: {ex.Message}");
            }
        }

        private async void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            var main = (Owner as MainWindow)!;
            if (main?.RecorderMgr.Evidence == null)
            {
                MessageBox.Show("エクスポートするデータがありません。");
                return;
            }
            var evidence = main.RecorderMgr.Evidence;

            try
            {
                foreach (var bm in main.RecorderMgr.Evidence.Bookmarks)
                {
                    main.VideoPlayer.Position = TimeSpan.FromSeconds(bm.Seconds);
                    await Task.Delay(500);
                    var scale = GetSelectedScale();
                    var savedPath = main.RecorderMgr.SaveSingleBookmarkImage(bm, main.VideoPlayer, scale);
                    bm.ImagePath = savedPath;
                }

                var htmlPath = ExportToHtml();
                var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + ".pdf";
                var filePath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                    main.RecorderMgr.CurrentFolder : OutputPathBox.Text, fileName);

                await ExportToPdfAsync(htmlPath, filePath);
                MessageBox.Show("PDF出力");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"出力エラー: {ex.Message}");
            }
        }

        private string ExportToHtml(bool isOriginalSize = false)
        {
            var main = (Owner as MainWindow)!;
            var evidence = main.RecorderMgr.Evidence;

            string fileName = Path.GetFileNameWithoutExtension(evidence?.VideoFileName) + ".html";
            string htmlPath = Path.Combine(main.RecorderMgr.CurrentFolder, fileName);
            var sb = new StringBuilder();

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

            return htmlPath;
        }

        private async Task ExportToPdfAsync(string htmlPath, string pdfPath)
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

        private void ExportToExcel()
        {
            var main = (Owner as MainWindow)!;
            var evidence = main.RecorderMgr.Evidence;

            var fileName = Path.GetFileNameWithoutExtension(evidence?.VideoFileName) + ".xlsx";
            var fullPath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                main.RecorderMgr.CurrentFolder : OutputPathBox.Text, fileName);

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
    }
}
