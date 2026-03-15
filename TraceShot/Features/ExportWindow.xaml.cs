
namespace TraceShot.Features
{
    using ClosedXML.Excel;
    using ClosedXML.Excel.Drawings;
    using DocumentFormat.OpenXml.Drawing.Charts;
    using Microsoft.Win32;
    using PuppeteerSharp;
    using PuppeteerSharp.Media;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using TraceShot.Models;
    using TraceShot.Services;
    using TraceShot.ViewModels;
    using MessageBox = System.Windows.MessageBox;
    using Path = System.IO.Path;

    /// <summary>
    /// ExportWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ExportWindow : Window
    {
        // クラスのメンバとしてリストを保持
        private ObservableCollection<ExportItemViewModel> _exportItems = new ObservableCollection<ExportItemViewModel>();

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
            LoadingOverlay.Visibility = Visibility.Visible;
            ExportProgressBar.Value = 0;

            var progress = new Progress<int>(value => {
                Dispatcher.BeginInvoke(() => {
                    ExportProgressBar.Value = value;
                    StatusText.Text = $"書き出し中... {value}%";
                }, System.Windows.Threading.DispatcherPriority.Render);
            });

            try
            {
                await exportAction(progress);
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

        private void ExportHtml_Click(object sender, RoutedEventArgs e)
        {
            var evidence = RecService.Instance.Evidence;
            var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + "_full.html";
            var fullPath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                RecService.Instance.CurrentFolder : OutputPathBox.Text, fileName);

            try
            {
                CaptureGroupBox.IsEnabled = false;
                OutputGroupBox.IsEnabled = false;
                var sb = GenerateHtmlFile();
                File.WriteAllText(fullPath, sb.ToString());

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(fullPath) { UseShellExecute = true }
                };
                p.Start();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CaptureGroupBox.IsEnabled = true;
                OutputGroupBox.IsEnabled = true;
            }

        }

        private string GenerateHtmlFile()
        {
            var targetItems = GetTargetItems();
            if (targetItems.Count == 0)
            {
                throw new Exception("出力対象が選択されていません。");
            }

            var showRelative = RadioRelativeTime.IsChecked == true;
            var timeHeader = showRelative ? "経過時間" : "実行時間";

            var sb = new StringBuilder();
            // --- HTML ヘッダー ---
            sb.AppendLine("<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>");
            sb.AppendLine("<title>エビデンス報告書</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: sans-serif; margin: 20px; background: #f0f2f5; }");
            sb.AppendLine(".container { max-width: 98%; margin: auto; background: white; padding: 20px; box-shadow: 0 0 15px rgba(0,0,0,0.1); }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; }");
            sb.AppendLine(".col-time { width: 100px; text-align: center; }");
            sb.AppendLine(".col-note { width: 150px; }");
            sb.AppendLine("th, td { border: 1px solid #dee2e6; padding: 10px; vertical-align: top; }");
            sb.AppendLine("th { background-color: #4472C4; color: white; }");

            // --- 画像とモーダルのスタイル ---
            sb.AppendLine(".modal { display: none; position: fixed; z-index: 999; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.9); cursor: zoom-out; }");
            sb.AppendLine(".modal-content { margin: auto; display: block; max-width: 95%; max-height: 95%; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); border: 2px solid #fff; }");
            sb.AppendLine(".ss-image { max-width: 100%; max-height: 500px; width: auto; height: auto; display: block; margin: 0 auto; cursor: zoom-in; transition: 0.2s; object-fit: contain; }");
            sb.AppendLine(".ss-image:hover { opacity: 0.8; transform: scale(1.01); }");
            sb.AppendLine(".col-ss { background-color: #f8f9fa; text-align: center; }");
            sb.AppendLine("</style></head><body><div class='container'>");

            sb.AppendLine("<table>");
            sb.AppendLine($"<thead><tr><th class='col-time'>{timeHeader}</th><th class='col-note'>コメント</th><th>スクリーンショット</th></tr></thead><tbody>");

            var startAt = RecService.Instance.Evidence.RecordingDate;
            foreach (var item in targetItems)
            {
                string time = showRelative ?
                    $"+{(int)item.Time.TotalHours:D2}:{item.Time.Minutes:D2}:{item.Time.Seconds:D2}"
                    : startAt.Add(item.Time).ToString("yyyy/MM/dd<br/>HH:mm:ss");
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td class='col-time'>{time}</td>");
                sb.AppendLine($"<td class='col-note'>{item.OriginalBookmark.Note}</td>");

                string base64String = "data:image/png;base64," + BitmapSourceToBase64(item.SnapshotImage);
                sb.AppendLine($"<td class='col-ss'><img src='{base64String}' class='ss-image' onclick='openModal(this)'></td>");
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</tbody></table></div>");

            // --- モーダル要素 ---
            sb.AppendLine("<div id='imgModal' class='modal' onclick='this.style.display=\"none\"'>");
            sb.AppendLine("  <img id='fullImg' class='modal-content'>");
            sb.AppendLine("</div>");

            // --- キーボードナビゲーション対応スクリプト ---
            sb.AppendLine("<script>");
            sb.AppendLine("let currentImgIndex = 0;");
            sb.AppendLine("const allImages = [];");

            // 読み込み時に全画像のリストを作成
            sb.AppendLine("window.onload = () => {");
            sb.AppendLine("  document.querySelectorAll('.ss-image').forEach((img, index) => {");
            sb.AppendLine("    allImages.push(img.src);");
            sb.AppendLine("    img.setAttribute('data-index', index);");
            sb.AppendLine("  });");
            sb.AppendLine("};");

            // モーダルを開く
            sb.AppendLine("function openModal(imgElement) {");
            sb.AppendLine("  const modal = document.getElementById('imgModal');");
            sb.AppendLine("  currentImgIndex = parseInt(imgElement.getAttribute('data-index'));");
            sb.AppendLine("  updateModalImage();");
            sb.AppendLine("  modal.style.display = 'block';");
            sb.AppendLine("}");

            // 画像の更新
            sb.AppendLine("function updateModalImage() {");
            sb.AppendLine("  document.getElementById('fullImg').src = allImages[currentImgIndex];");
            sb.AppendLine("}");

            // キーボードイベント
            sb.AppendLine("document.onkeydown = function(e) {");
            sb.AppendLine("  const modal = document.getElementById('imgModal');");
            sb.AppendLine("  if (modal.style.display !== 'block') return;");
            sb.AppendLine("  if (e.key === 'ArrowRight') { currentImgIndex = (currentImgIndex + 1) % allImages.length; updateModalImage(); }");
            sb.AppendLine("  else if (e.key === 'ArrowLeft') { currentImgIndex = (currentImgIndex - 1 + allImages.length) % allImages.length; updateModalImage(); }");
            sb.AppendLine("  else if (e.key === 'Escape') { modal.style.display = 'none'; }");
            sb.AppendLine("};");
            sb.AppendLine("</script>");

            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private async void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            var evidence = RecService.Instance.Evidence;
            var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + ".pdf";
            var filePath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                RecService.Instance.CurrentFolder : OutputPathBox.Text, fileName);

            try
            {
                CaptureGroupBox.IsEnabled = false;
                OutputGroupBox.IsEnabled = false;
                string htmlContent = htmlContent = GenerateHtmlFileForPdf();
                string tempHtmlPath = Path.Combine(Path.GetTempPath(), "TraceShot_temp.html");

                File.WriteAllText(tempHtmlPath, htmlContent);
                await ExportToPdfAsync(tempHtmlPath, filePath);
                if (File.Exists(tempHtmlPath)) File.Delete(tempHtmlPath);
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(filePath) { UseShellExecute = true }
                };
                p.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CaptureGroupBox.IsEnabled = true;
                OutputGroupBox.IsEnabled = true;
            }
        }

        private string GenerateHtmlFileForPdf()
        {
            var targetItems = GetTargetItems();
            if (targetItems.Count == 0)
            {
                throw new Exception("出力対象が選択されていません。");
            }

            var showRelative = RadioRelativeTime.IsChecked == true;
            var timeHeader = showRelative ? "経過時間" : "実行時間";

            var sb = new StringBuilder();
            // --- PDF用L ヘッダー ---
            sb.AppendLine("<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>");
            sb.AppendLine("<title>エビデンス報告書</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: sans-serif; margin: 0; padding: 0; background: #f0f2f5; }");
            sb.AppendLine(".container { width: 100%; margin: 0; background: white; padding: 10px; box-sizing: border-box; }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; table-layout: fixed; }"); // table-layout: fixed が重要
            sb.AppendLine(".col-time { width: 100px; text-align: center; font-size: 0.9em; }");
            sb.AppendLine(".col-note { width: 150px; word-wrap: break-word; }");
            sb.AppendLine(".col-ss { width: auto; }");
            sb.AppendLine("th, td { border: 1px solid #dee2e6; padding: 8px; vertical-align: top; }");
            sb.AppendLine("th { background-color: #4472C4; color: white; }");
            sb.AppendLine(".ss-image { width: 100%; height: auto; display: block; }");
            sb.AppendLine("@media print {");
            sb.AppendLine("  body { background: white; }");
            sb.AppendLine("  .container { padding: 0; box-shadow: none; }");
            sb.AppendLine("  th { -webkit-print-color-adjust: exact; }");
            sb.AppendLine("}");
            sb.AppendLine("</style></head><body><div class='container'>");
            sb.AppendLine("<table>");
            sb.AppendLine($"<thead><tr><th class='col-time'>{timeHeader}</th><th class='col-note'>コメント</th><th class='col-ss'>スクリーンショット</th></tr></thead>");
            sb.AppendLine("<tbody>");
            var startAt = RecService.Instance.Evidence.RecordingDate;
            foreach (var item in targetItems)
            {
                string time = showRelative ?
                    $"+{(int)item.Time.TotalHours:D2}:{item.Time.Minutes:D2}:{item.Time.Seconds:D2}"
                    : startAt.Add(item.Time).ToString("yyyy/MM/dd<br/>HH:mm:ss");

                sb.AppendLine("<tr>");
                sb.AppendLine($"<td class='col-time'>{time}</td>");
                sb.AppendLine($"<td class='col-note'>{item.Note}</td>");

                string base64String = "data:image/png;base64," + BitmapSourceToBase64(item.SnapshotImage);
                sb.AppendLine($"<td class='col-ss'><img src='{base64String}' class='ss-image'></td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</table></div></body></html>");
            return sb.ToString();
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

                    // 画面表示モード
                    await page.EmulateMediaTypeAsync(MediaType.Screen);

                    // 💡 念のため、画像などの読み込み完了を少し待つ
                    await page.WaitForNetworkIdleAsync();

                    // 4. PDFとして保存
                    await page.PdfAsync(pdfPath, new PdfOptions
                    {
                        Format = PaperFormat.A4,
                        PrintBackground = true, // 背景色やヘッダーの色を出す
                    });
                }
            }
        }

        private async void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var evidence = RecService.Instance.Evidence;
            var fileName = Path.GetFileNameWithoutExtension(evidence.VideoFileName) + ".xlsx";
            var fullPath = Path.Combine(string.IsNullOrEmpty(OutputPathBox.Text) ?
                RecService.Instance.CurrentFolder : OutputPathBox.Text, fileName);

            try
            {
                CaptureGroupBox.IsEnabled = false;
                OutputGroupBox.IsEnabled = false;

                await Task.Run(() =>
                {
                    SaveAsExcel(fullPath);
                    if (File.Exists(fullPath))
                    {
                        string argument = $"/select,\"{fullPath}\"";
                        Process.Start("explorer.exe", argument);
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CaptureGroupBox.IsEnabled = true;
                OutputGroupBox.IsEnabled = true;
            }
        }

        private void SaveAsExcel(string fullPath)
        {
            // UI情報の取得（前述の通りInvokeで）
            bool showRelative = false;
            double scale = 1.0;
            this.Dispatcher.Invoke(() => {
                showRelative = RadioRelativeTime.IsChecked == true;
                scale = GetSelectedScale();
            });

            var targetItems = GetTargetItems();
            if (targetItems.Count == 0) return;

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("録画レポート"); // 全画像をこのシートに集約

            // デザインの基本設定
            ws.Column(1).Width = 15;
            ws.Column(2).Width = 60;

            int currentAnchorRow = 1; // 書き出し開始行

            foreach (var item in targetItems)
            {
                // --- 1. テキスト情報の配置 ---
                // ヘッダー（時間）
                var timeCell = ws.Cell(currentAnchorRow, 1);
                timeCell.Value = showRelative ? "経過時間" : "実行時間";

                var valCell = ws.Cell(currentAnchorRow, 2);
                if (showRelative)
                    valCell.Value = item.Time.ToString(@"hh\:mm\:ss");
                else
                    valCell.Value = RecService.Instance.Evidence.RecordingDate.Add(item.Time).ToString("yyyy/MM/dd HH:mm:ss");

                // コメント
                ws.Cell(currentAnchorRow + 1, 1).Value = "コメント";
                ws.Cell(currentAnchorRow + 1, 2).Value = item.Note;

                // スタイル適用
                var headerRange = ws.Range(currentAnchorRow, 1, currentAnchorRow + 1, 1);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4");
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;

                // --- 2. 画像の挿入 ---
                if (item.SnapshotImage != null)
                {
                    var imgData = ImageData.GetImageData(item.SnapshotImage);
                    if (imgData != null)
                    {
                        using var imageStream = new MemoryStream(imgData.Bytes);
                        var picture = ws.AddPicture(imageStream);

                        picture.Placement = XLPicturePlacement.Move;

                        // ⭐ 数値指定ではなく、Scale(1.0) を基準にする
                        // これにより、Excelが「画像本来のサイズ」として配置しようとします
                        double dpiScale = VisualTreeHelper.GetDpi(this).DpiScaleX;
                        picture.Width = (int)Math.Round((imgData.PixelWidth * item.Scale) / dpiScale);
                        picture.Height = (int)Math.Round((imgData.PixelHeight * item.Scale) / dpiScale);

                        int imageStartRow = currentAnchorRow + 3;
                        picture.MoveTo(ws.Cell(imageStartRow, 1));

                        // 高さは配置後の picture オブジェクトから自動取得
                        int rowsUsedByImage = (int)Math.Ceiling(picture.Height / 20.0) + 2;
                        currentAnchorRow = imageStartRow + rowsUsedByImage + 1;
                    }
                }
                else
                {
                    // 画像がない場合はテキスト分だけ進める
                    currentAnchorRow += 2;
                }

                // 区切り線代わりに少し余白
                currentAnchorRow++;
            }

            workbook.SaveAs(fullPath);
        }


        private async void StartCapture_Click(object sender, RoutedEventArgs e)
        {
            _exportItems.Clear();
            ExportPreviewList.ItemsSource = _exportItems;

            //var marks = RecService.Instance.Bookmarks;
            var main = Owner as MainWindow;
            var scale = GetSelectedScale();
            int index = 0;
            await RunExportTask(async (progress) =>
            {
                //for (int i = 0; i < RecService.Instance.Bookmarks.Count; i++)
                foreach(var bm in RecService.Instance.Bookmarks)
                {
                    // 1. 指定時間に移動
                    main.VideoPlayer.Position = bm.Time;
                    await Task.Delay(500); // 描画待ち

                    // 2. キャプチャ実行
                    var snapshot = new VideoSnapshotInfo(main.VideoPlayer);
                    var result = RecService.Instance.SaveImage(bm, snapshot, scale);

                    // 3. ViewModelを作成してリストに追加
                    // UIスレッドで実行する必要があるため、Dispatcher経由で行う
                    this.Dispatcher.Invoke(() => {
                        _exportItems.Add(new ExportItemViewModel
                        {
                            OriginalBookmark = bm,
                            SnapshotImage = result?.Bitmap,
                            ImagePath = result?.Path,
                            Order = index++,
                            Scale = scale,
                        });
                    });

                    progress.Report((index) * 100 / RecService.Instance.Bookmarks.Count);
                }

                // 全撮影終了後のメッセージ
                StatusText.Text = "撮影完了。出力する画像を選択・確認してください。";
                OutputGroupBox.IsEnabled = true;
                EmptyGuideText.Visibility = Visibility.Collapsed;
            });

            this.Dispatcher.Invoke(() =>
            {
                // _exportItems を時系列（Time）で並べ替えたリストを作成
                var sortedList = _exportItems.OrderBy(x => x.Time).ToList();

                // 元の ObservableCollection をクリアして、正しい順序で再登録
                _exportItems.Clear();
                foreach (var item in sortedList)
                {
                    _exportItems.Add(item);
                }

                // ガイド表示を更新
                StatusText.Text = "撮影完了。並べ替えや選択が可能です。";
            });
        }

        // ドラッグ開始の処理
        private void Item_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && sender is FrameworkElement element)
            {
                // ドラッグするデータ（ViewModel）を取得
                var data = element.DataContext as ExportItemViewModel;
                if (data == null) return;

                // ドラッグ＆ドロップ操作を開始
                DragDrop.DoDragDrop(element, data, DragDropEffects.Move);
            }
        }

        // ドロップ時の処理
        private void Item_Drop(object sender, DragEventArgs e)
        {
            // ドロップされたデータ（移動元）
            var sourceData = e.Data.GetData(typeof(ExportItemViewModel)) as ExportItemViewModel;
            // ドロップ先のデータ（移動先）
            var targetData = (sender as FrameworkElement)?.DataContext as ExportItemViewModel;

            if (sourceData != null && targetData != null && sourceData != targetData)
            {
                // コレクション内でのインデックスを取得
                int oldIndex = _exportItems.IndexOf(sourceData);
                int newIndex = _exportItems.IndexOf(targetData);

                if (oldIndex >= 0 && newIndex >= 0)
                {
                    // ObservableCollection の要素を移動（UIも自動更新される）
                    _exportItems.Move(oldIndex, newIndex);

                    // 必要に応じて全アイテムの Order プロパティを更新
                    for (int i = 0; i < _exportItems.Count; i++)
                    {
                        _exportItems[i].Order = i;
                    }
                }
            }
        }

        private List<ExportItemViewModel> GetTargetItems()
        {
            return _exportItems.Where(x => x.IsSelected).ToList();
        }

        public string BitmapSourceToBase64(BitmapSource bitmapSource)
        {
            if (bitmapSource == null) return string.Empty;

            // 1. エンコーダーを準備（PNGまたはJPEG）
            var encoder = new PngBitmapEncoder(); // 画質重視ならPNG、容量重視ならJpegBitmapEncoder
            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));

            using (var ms = new MemoryStream())
            {
                // 2. メモリストリームに保存
                encoder.Save(ms);
                byte[] imageBytes = ms.ToArray();

                // 3. バイト配列をBase64文字列に変換
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
}
