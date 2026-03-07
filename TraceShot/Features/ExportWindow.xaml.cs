using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using Microsoft.Win32;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using TraceShot.Models;
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
            OutputPathBox.Text = Properties.Settings.Default.SavePath;

            // 💡 画面が開いた時に、現在の動画の見た目をプレビューにセット
            this.Loaded += (s, e) => {
                var main = Owner as MainWindow;
                if (main != null)
                {
                    // 現在の表示内容をキャプチャしてプレビューに表示するロジック（RenderTargetBitmap等）
                    // または、最新のブックマーク画像を一時的に表示
                }
            };
        }

        private async Task RunExportTask(Func<IProgress<int>, Task> exportAction)
        {
            // 1. UIをロックしてプログレス表示
            LoadingOverlay.Visibility = Visibility.Visible;
            ExportProgressBar.Value = 0;

            // 進捗報告用のインスタンス
            var progress = new Progress<int>(value => {
                ExportProgressBar.Value = value;
            });

            try
            {
                // 2. 重い処理を実行
                await exportAction(progress);
                //MessageBox.Show("エクスポートが完了しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}");
            }
            finally
            {
                // 3. UIを復帰
                ExportProgressBar.Value = 0;
                StatusText.Text = "準備完了";
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
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

        private async void ExportHtml_Click(object sender, RoutedEventArgs e)
        {
            var main = Owner as MainWindow;
            if (main?.RecorderMgr.Evidence == null) return;

            var evidence = main.RecorderMgr.Evidence;
            var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + "_full.html";
            var fullPath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                main.RecorderMgr.CurrentFolder : OutputPathBox.Text, fileName);
            var crop = evidence.Bookmarks.SelectMany(b => b.MarkRects).FirstOrDefault(r => r.IsCropArea);

            await RunExportTask(async (progress) =>
            {
                var marks = main.RecorderMgr.Evidence.Bookmarks;
                int total = marks.Count;
                var scale = GetSelectedScale();

                for (int i = 0; i < total; i++)
                {
                    var bm = marks[i];
                    Dispatcher.Invoke(() => StatusText.Text = $"画像生成中... ({i + 1}/{total})");
                    await Dispatcher.InvokeAsync(async () => {
                        main.VideoPlayer.Position = bm.Time;
                    });
                    await Task.Delay(100);
                    await Dispatcher.InvokeAsync(() => {
                        (string? Path, BitmapSource? Bitmap)? result;
                        if (crop != null)
                        {
                            result = main.RecorderMgr.SaveCroppedBookmarkImage(bm, main.VideoPlayer, crop, scale); ;
                        }
                        else
                        {
                            result = main.RecorderMgr.SaveSingleBookmarkImage(bm, main.VideoPlayer, scale);
                        }

                        bm.ImagePath = result?.Path;
                        if (result?.Bitmap != null)
                        {
                            PreviewImage.Source = result?.Bitmap;
                        }
                    });
                    progress.Report((i + 1) * 100 / (total + 1));
                }
                Dispatcher.Invoke(() => StatusText.Text = "レポート出力中...");

                await Task.Run(() =>
                {
                    var htmlContent = GenerateHtmlFile(evidence);
                    File.WriteAllText(fullPath, htmlContent);
                    var p = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo(fullPath) { UseShellExecute = true }
                    };
                    p.Start();
                });
                progress.Report(100);
            });
        }

        private string GenerateHtmlFile(RecordingEvidence evidence)
        {
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
            return sb.ToString();
        }

        private async void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            var main = Owner as MainWindow;
            if (main?.RecorderMgr.Evidence == null) return;

            var evidence = main.RecorderMgr.Evidence;
            // 出力先の決定
            var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + ".pdf";
            var filePath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                main.RecorderMgr.CurrentFolder : OutputPathBox.Text, fileName);
            var crop = evidence.Bookmarks.SelectMany(b => b.MarkRects).FirstOrDefault(r => r.IsCropArea);

            await RunExportTask(async (progress) =>
            {
                var marks = evidence.Bookmarks;
                int total = marks.Count;
                var scale = GetSelectedScale();

                // 1. 全ての画像を生成（HTML時と同じロジック）
                for (int i = 0; i < total; i++)
                {
                    var bm = marks[i];
                    Dispatcher.Invoke(() => StatusText.Text = $"画像生成中... ({i + 1}/{total})");
                    await Dispatcher.InvokeAsync(async () => {
                        main.VideoPlayer.Position = bm.Time;
                    });
                    await Task.Delay(100);
                    await Dispatcher.InvokeAsync(() => {
                        (string? Path, BitmapSource? Bitmap)? result;
                        if (crop != null)
                        {
                            result = main.RecorderMgr.SaveCroppedBookmarkImage(bm, main.VideoPlayer, crop, scale); ;
                        }
                        else
                        {
                            result = main.RecorderMgr.SaveSingleBookmarkImage(bm, main.VideoPlayer, scale);
                        }
                        bm.ImagePath = result?.Path;
                        if (result?.Bitmap != null)
                        {
                            PreviewImage.Source = result?.Bitmap;
                        }
                    });
                    progress.Report((i + 1) * 100 / (total + 2)); // 最後の2ステップをHTML/PDF用に残す
                }

                // 2. HTML 文字列の生成（既存の GenerateHtmlFile を活用）
                Dispatcher.Invoke(() => StatusText.Text = "一時レポートを作成中...");
                string htmlContent = "";
                await Task.Run(() => {
                    htmlContent = GenerateHtmlFile(evidence);
                });
                progress.Report((total + 1) * 100 / (total + 2));

                // 3. 一時的なHTMLファイルを保存してPDFに変換
                Dispatcher.Invoke(() => StatusText.Text = "PDFに変換中（ブラウザ起動）...");
                await Task.Run(async () =>
                {
                    // PDF変換用に一時ファイルとして書き出す
                    string tempHtmlPath = Path.Combine(Path.GetTempPath(), "TraceShot_temp.html");
                    File.WriteAllText(tempHtmlPath, htmlContent);

                    // PuppeteerSharp を使った変換処理
                    await ExportToPdfAsync(tempHtmlPath, filePath);

                    // 使い終わった一時ファイルを削除（任意）
                    if (File.Exists(tempHtmlPath)) File.Delete(tempHtmlPath);

                    // 完了後にPDFを開く
                    var p = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true }
                    };
                    p.Start();
                });

                progress.Report(100);
            });
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

        private async void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var main = Owner as MainWindow;
            if (main?.RecorderMgr.Evidence == null) return;

            var evidence = main.RecorderMgr.Evidence;
            var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + ".xlsx";
            var fullPath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                main.RecorderMgr.CurrentFolder : OutputPathBox.Text, fileName);
            var crop = evidence.Bookmarks.SelectMany(b => b.MarkRects).FirstOrDefault(r => r.IsCropArea);

            await RunExportTask(async (progress) =>
            {
                var marks = evidence.Bookmarks;
                int total = marks.Count;
                var scale = GetSelectedScale();

                // 1. 全ての画像を生成（HTML/PDFと同じ安定化ロジック）
                for (int i = 0; i < total; i++)
                {
                    var bm = marks[i];
                    Dispatcher.Invoke(() => StatusText.Text = $"画像生成中... ({i + 1}/{total})");
                    await Dispatcher.InvokeAsync(async () => {
                        main.VideoPlayer.Position = bm.Time;
                    });
                    await Task.Delay(100);
                    await Dispatcher.InvokeAsync(() => {
                        (string? Path, BitmapSource? Bitmap)? result;
                        if (crop != null)
                        {
                            result = main.RecorderMgr.SaveCroppedBookmarkImage(bm, main.VideoPlayer, crop, scale); ;
                        }
                        else
                        {
                            result = main.RecorderMgr.SaveSingleBookmarkImage(bm, main.VideoPlayer, scale);
                        }
                        bm.ImagePath = result?.Path;
                        if (result?.Bitmap != null)
                        {
                            PreviewImage.Source = result?.Bitmap;
                        }
                    });
                    progress.Report((i + 1) * 100 / (total + 1));
                }

                // 2. Excelファイルの生成
                Dispatcher.Invoke(() => StatusText.Text = "Excelファイルを構築中...");
                await Task.Run(() =>
                {
                    // 💡 既存の ExportToExcel() メソッドを呼び出す
                    // 引数で fullPath を渡せるように修正しておくとスムーズです
                    SaveAsExcel(evidence, fullPath, scale);

                    if (File.Exists(fullPath))
                    {
                        string argument = $"/select,\"{fullPath}\"";
                        System.Diagnostics.Process.Start("explorer.exe", argument);
                    }
                });

                progress.Report(100);
            });
        }
        private void SaveAsExcel(RecordingEvidence evidence, string fullPath, double selectedScale)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                for (int i = 0; i < evidence.Bookmarks.Count; i++)
                {
                    var bm = evidence.Bookmarks[i];
                    if (bm == null) continue;
                    // ※シート名に使えない記号を置換します
                    string sheetName = $"{i + 1}_{bm.Time:HH-mm-ss_fff}";
                    var ws = workbook.Worksheets.Add(sheetName);

                    // 2. テキスト情報の配置
                    ws.Cell(1, 1).Value = "経過時間";
                    ws.Cell(1, 2).Value = bm.Time;
                    ws.Cell(2, 1).Value = "コメント";
                    ws.Cell(2, 2).Value = bm.Note;

                    // デザイン調整
                    var headerRange = ws.Range("A1:A2");
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4");
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Font.Bold = true;
                    ws.Column(1).Width = 15;
                    ws.Column(2).Width = 50;

                    // 3. 画像の挿入（指定倍率を適用）
                    if (!string.IsNullOrEmpty(bm.ImagePath) && File.Exists(bm.ImagePath))
                    {
                        var picture = ws.AddPicture(bm.ImagePath);

                        // 💡 修正：スケールを変更する前に配置モードを設定する
                        picture.Placement = XLPicturePlacement.Move;

                        // ComboBoxで選ばれた倍率（0.25〜1.0）を適用
                        picture.Scale(selectedScale);

                        // 4行目から画像を配置（テキストと重ならないように）
                        picture.MoveTo(ws.Cell(4, 1), 5, 5);

                        // 画像が隠れないように、配置した場所の行高さを調整
                        // Excelの行高さ制限（409.5）に注意しながら設定
                        double rowHeight = (picture.Height * 0.75) + 20;
                        ws.Row(4).Height = Math.Min(rowHeight, 409);
                    }
                }

                // 最後にブックを保存
                workbook.SaveAs(fullPath);
            }
        }
    }
}
