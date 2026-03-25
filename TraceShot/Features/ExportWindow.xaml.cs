
namespace TraceShot.Features
{
    using ClosedXML.Excel;
    using ClosedXML.Excel.Drawings;
    using CommunityToolkit.Mvvm.ComponentModel;
    using Microsoft.Win32;
    using PuppeteerSharp;
    using PuppeteerSharp.Media;
    using System.Collections.ObjectModel;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Media.Animation;
    using TraceShot.Models;
    using TraceShot.Services;
    using TraceShot.ViewModels;
    using Border = System.Windows.Controls.Border;
    using MessageBox = System.Windows.MessageBox;
    using Path = System.IO.Path;

    [INotifyPropertyChanged]
    public partial class ExportWindow : Window
    {
        public ObservableCollection<ExportItemViewModel> ExportItems { get; set; } = [];
        private bool _isPreviewMode = false;
        private int _currentIdx = 0;

        private readonly ExportCacheManager _cacheManager;

        [ObservableProperty] private ExportItemViewModel? _selectedPreviewItem;

        [ObservableProperty] private string? _previewCounterText;

        public ExportWindow(ExportCacheManager cacheManager)
        {
            InitializeComponent();
            _cacheManager = cacheManager;

            double initialScale = _cacheManager.LastScale > 0 ? _cacheManager.LastScale : 0.75;
            SetScaleComboBoxValue(initialScale);

            var viewModels = new List<ExportItemViewModel>();
            var sortedEntries = RecService.Instance.Entries
                .OrderBy(b => b.ExportOrder)
                .ThenBy(b => b.Time);

            foreach (var b in sortedEntries)
            {
                var cachedImage = _cacheManager.GetCachedImage(b.Id, initialScale);
                var vm = new ExportItemViewModel(b, cachedImage, _cacheManager)
                {
                    IsSelected = b.IsExportEnabled,
                    Order = b.ExportOrder,
                };
                vm.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(ExportItemViewModel.IsSelected))
                    {
                        RefreshExportIds();
                    }
                };
                viewModels.Add(vm);
            }

            ExportItems = new ObservableCollection<ExportItemViewModel>(viewModels);
            RefreshExportIds();

            for (int i = 0; i < ExportItems.Count; i++)
            {
                ExportItems[i].SerialNumber = (i + 1).ToString("D2");
            }
            ExportItems.CollectionChanged += (s, e) =>
            {
                if (e.Action == NotifyCollectionChangedAction.Move)
                {
                    // ★ 1. 連番を上から順に振り直す
                    for (int i = 0; i < ExportItems.Count; i++)
                    {
                        // 1始まりで、2桁のゼロ埋め（01, 02...）にする
                        ExportItems[i].SerialNumber = (i + 1).ToString("D2");

                        // 2. ついでに Bookmark の Order も更新（前回の永続化ロジック）
                        ExportItems[i].OriginalBookmark.ExportOrder = i;
                    }

                    // 3. JSONを上書き保存
                    RecService.Instance.SaveEvidenceJson();
                }
            };
            bool hasImages = ExportItems.Any(item => item.SnapshotImage != null);

            EmptyGuideText.Visibility = hasImages ? Visibility.Collapsed : Visibility.Visible;
            ExportPreviewList.Visibility = !hasImages ? Visibility.Collapsed : Visibility.Visible;

            DataContext = this;

            OutputPathBox.Text = Properties.Settings.Default.SavePath;
        }

        private void SetScaleComboBoxValue(double scale)
        {
            foreach (var item in ExportScaleComboBox.Items.OfType<ComboBoxItem>())
            {
                if (double.TryParse(item.Tag?.ToString(), out double val))
                {
                    if (Math.Abs(val - scale) < 0.001)
                    {
                        ExportScaleComboBox.SelectedItem = item;
                        return;
                    }
                }
            }
            ExportScaleComboBox.SelectedIndex = 0; // デフォルト
        }

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.DataContext is ExportItemViewModel vm)
            {
                // 操作された瞬間にマネージャーに覚えさせる
                _cacheManager.UpdateSelection(vm.OriginalBookmark.Id, vm.IsSelected);
            }
        }

        private async Task RunExportTask(Func<IProgress<int>, Task> exportAction)
        {
            ExportPreviewList.Visibility = Visibility.Collapsed;
            EmptyGuideText.Visibility = Visibility.Collapsed;
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
                ExportPreviewList.Visibility = Visibility.Visible;

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
            if (targetItems.Count == 0) throw new Exception("出力対象が選択されていません。");

            var startAt = RecService.Instance.Evidence.RecordingDate;
            var showRelative = RadioRelativeTime.IsChecked ?? false;

            var sb = new StringBuilder();
            // --- HTML ヘッダー ---
            sb.AppendLine("<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>");
            sb.AppendLine("<title>エビデンス報告書</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: sans-serif; margin: 20px; background: #f0f2f5; }");
            sb.AppendLine(".container { max-width: 98%; margin: auto; background: white; padding: 20px; box-shadow: 0 0 15px rgba(0,0,0,0.1); }");

            // --- サマリ用のスタイル追加 ---
            sb.AppendLine(".summary-section { margin-bottom: 40px; border: 1px solid #4472C4; border-radius: 8px; overflow: hidden; }");
            sb.AppendLine(".summary-header { background: #4472C4; color: white; padding: 10px 15px; font-weight: bold; font-size: 1.2em; }");
            sb.AppendLine(".summary-table { width: 100%; border-collapse: collapse; }");
            sb.AppendLine(".summary-table th { background: #f8f9fa; color: #333; }");

            // --- CSS部分に追加 ---
            sb.AppendLine(".res-ok { color: #28a745; font-weight: bold; }");   // 緑
            sb.AppendLine(".res-ng { color: #dc3545; font-weight: bold; }");   // 赤
            sb.AppendLine(".res-pend { color: #fd7e14; font-weight: bold; }"); // オレンジ

            sb.AppendLine(".table-wrapper { border: 1px solid #dee2e6; border-radius: 10px; overflow: hidden; margin-top: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; border: none; }");

            // ケース番号と時間の入る列
            sb.AppendLine(".col-case { width: 110px; text-align: center; background-color: #f8f9fa; }");
            sb.AppendLine(".case-id { font-weight: bold; color: #4472C4; font-size: 1.1em; display: block; }");
            sb.AppendLine(".step-id { font-size: 0.9em; color: #555; display: block; margin: 2px 0; }");
            sb.AppendLine(".case-time { font-size: 0.85em; color: #666; display: block; margin-top: 4px; }");
            sb.AppendLine(".col-note { width: 150px; }");
            sb.AppendLine("th, td { border: 1px solid #dee2e6; padding: 12px 10px; vertical-align: top; }");
            sb.AppendLine("th { background-color: #4472C4; color: white; border-top: none; }");
            sb.AppendLine("tbody tr:last-child td { border-bottom: none; }");

            // --- 画像とモーダルのスタイル ---
            sb.AppendLine(".modal { display: none; position: fixed; z-index: 999; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.9); cursor: zoom-out; }");
            sb.AppendLine(".modal-content { margin: auto; display: block; max-width: 95%; max-height: 95%; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); border: 2px solid #fff; }");
            sb.AppendLine(".ss-image { max-width: 100%; max-height: 500px; width: auto; height: auto; display: block; margin: 0 auto; cursor: zoom-in; transition: 0.2s; object-fit: contain; }");
            sb.AppendLine(".ss-image:hover { opacity: 0.8; transform: scale(1.01); }");
            sb.AppendLine(".col-ss { background-color: #f8f9fa; text-align: center; }");
            sb.AppendLine("</style></head><body><div class='container'>");

            if (CheckIncludeSummary.IsChecked ?? false)
            {
                // --- 1. テストケースサマリの差し込み ---
                sb.AppendLine("<div class='summary-section'>");
                sb.AppendLine("<div class='summary-header'>テスト実施サマリ</div>");
                sb.AppendLine("<table class='summary-table'>");
                sb.AppendLine("<thead><tr><th>ケースNo</th><th>実施日時</th><th>開始時間</th><th>終了時間</th><th>ステップ数</th><th>結果</th></tr></thead><tbody>");

                // RecServiceからサマリを取得（StartTimeでソート済み）
                var summaries = RecService.Instance.Evidence.GetSummary();

                foreach (var s in summaries)
                {
                    if (s.FinalResult == TestResult.SS) continue;

                    string resultText = s.FinalResult.ToString(); // OK, NG, PEND が入る
                    string resultClass = s.FinalResult switch
                    {
                        TestResult.OK => "res-ok",
                        TestResult.NG => "res-ng",
                        TestResult.PEND => "res-pend",
                        _ => ""
                    };
                    // 日付、開始、終了をそれぞれ計算
                    string dateStr = startAt.Add(s.StartTime).ToString("yyyy/MM/dd");
                    string startStr = startAt.Add(s.StartTime).ToString("HH:mm:ss");
                    string endStr = startAt.Add(s.EndTime).ToString("HH:mm:ss");

                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td style='text-align:center;'>No. {s.CaseId}</td>");
                    sb.AppendLine($"<td style='text-align:center;'>{dateStr}</td>");
                    sb.AppendLine($"<td style='text-align:center;'>{startStr}</td>");
                    sb.AppendLine($"<td style='text-align:center;'>{endStr}</td>");
                    sb.AppendLine($"<td style='text-align:center;'>{s.StepCount} step(s)</td>");
                    sb.AppendLine($"<td style='text-align:center;' class='{resultClass}'>{resultText}</td>");
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</tbody></table></div>");
                sb.AppendLine("<h3>詳細ログ</h3>");
            }

            // --- 2. メインコンテンツ（詳細テーブル） ---
            sb.AppendLine("<div class='table-wrapper'>");
            sb.AppendLine("<table>");
            sb.AppendLine($"<thead><tr><th class='col-time'>ケース/Step</th><th class='col-note'>コメント</th><th>スクリーンショット</th></tr></thead><tbody>");

            var stepCounter = 0;
            var lastCaseId = -1;
            foreach (var item in targetItems)
            {
                var currentCaseId = item.OriginalBookmark.CaseId;
                if (lastCaseId != currentCaseId)
                {
                    lastCaseId = currentCaseId;
                    stepCounter = 1;
                }
                else
                {
                    stepCounter++;
                }

                string timeStr = showRelative ?
                    $"+{(int)item.Time.TotalHours:D2}:{item.Time.Minutes:D2}:{item.Time.Seconds:D2}"
                    : startAt.Add(item.Time).ToString("HH:mm:ss");

                sb.AppendLine("<tr>");

                // --- ケース番号と時間を併記した列 ---
                sb.AppendLine("  <td class='col-case'>");
                sb.AppendLine($"    <span class='case-id'>No {currentCaseId}</span>");
                sb.AppendLine($"    <span class='case-step'>Step {stepCounter}</span>");
                sb.AppendLine($"    <span class='case-time'>{timeStr}</span>");
                sb.AppendLine("  </td>");

                sb.AppendLine($"<td class='col-note'>{item.OriginalBookmark.Note}</td>");
                var base64Image = ImageService.BitmapSourceToBase64(item.SnapshotImage);
                var base64String = "data:image/png;base64," + base64Image;
                sb.AppendLine($"<td class='col-ss'><img src='{base64String}' class='ss-image' onclick='openModal(this)'></td>");
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</tbody></table></div></div>");

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
            if (targetItems.Count == 0) throw new Exception("出力対象が選択されていません。");

            var startAt = RecService.Instance.Evidence.RecordingDate;
            var showRelative = RadioRelativeTime.IsChecked ?? false;

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>");
            sb.AppendLine("<style>");

            // --- 基本設定 & 用紙設定 ---
            // @page で横向きを強制。marginを絞って描画領域を広げる
            sb.AppendLine("@page { size: A4 landscape; margin: 8mm; }");
            sb.AppendLine("body { font-family: sans-serif; margin: 0; padding: 0; background: white; width: 100%; }");
            sb.AppendLine(".container { width: 100%; margin: 0; padding: 0; box-sizing: border-box; }");

            // --- 1. サマリセクションのスタイル ---
            sb.AppendLine(".summary-section { margin-bottom: 30px; border: 1px solid #4472C4; border-radius: 8px; overflow: hidden; }");
            sb.AppendLine(".summary-header { background: #4472C4; color: white; padding: 8px 15px; font-weight: bold; }");
            sb.AppendLine(".summary-table { width: 100%; border-collapse: collapse; table-layout: fixed; }");
            // ヘッダーを薄いグレーに設定
            sb.AppendLine(".summary-table th { background: #f2f2f2; color: #333; border: 1px solid #dee2e6; padding: 8px; font-size: 0.9em; }");
            sb.AppendLine(".summary-table td { border: 1px solid #dee2e6; padding: 8px; text-align: center; font-size: 0.9em; }");
            sb.AppendLine(".res-ok { color: #28a745; font-weight: bold; }");
            sb.AppendLine(".res-ng { color: #dc3545; font-weight: bold; }");
            sb.AppendLine(".res-pend { color: #fd7e14; font-weight: bold; }");

            // --- 2. 詳細ログセクションのスタイル ---
            sb.AppendLine(".table-wrapper { border: 1px solid #dee2e6; border-radius: 8px; overflow: hidden; margin-top: 10px; }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; table-layout: fixed; word-wrap: break-word; }");
            // 各ページの冒頭にヘッダーを表示させるための設定
            sb.AppendLine("thead { display: table-header-group; }");
            sb.AppendLine("tfoot { display: table-footer-group; }"); // 必要であればフッターも同様

            // 詳細ヘッダーを濃い青に設定
            sb.AppendLine("th.col-case, th.col-note, th.col-ss { background-color: #4472C4; color: white; padding: 10px; -webkit-print-color-adjust: exact; }");

            // 横向き用の列幅配分（画像を最大化するために col-ss を 75% に）
            sb.AppendLine(".col-case { width: 10%; text-align: center; background-color: #f8f9fa; }");
            sb.AppendLine(".col-note { width: 15%; font-size: 0.9em; }");
            sb.AppendLine(".col-ss { width: 75%; text-align: center; background-color: #ffffff; }");
            sb.AppendLine("td { border: 1px solid #dee2e6; padding: 10px; vertical-align: top; }");

            // ケース/Step/時刻の縦並び
            sb.AppendLine(".case-id { font-weight: bold; color: #4472C4; display: block; font-size: 1.1em; }");
            sb.AppendLine(".step-id { font-size: 0.9em; display: block; margin: 2px 0; color: #333; }");
            sb.AppendLine(".case-time { font-size: 0.8em; color: #666; display: block; }");

            // スクリーンショットの最大化
            sb.AppendLine(".ss-image { width: 100%; height: auto; display: block; object-fit: contain; }");

            // 改ページ制御（行の途中で切れないようにする）
            //sb.AppendLine("tr { page-break-inside: avoid; }");
            sb.AppendLine("th { -webkit-print-color-adjust: exact; print-color-adjust: exact; }");
            sb.AppendLine("tr { page-break-inside: avoid; page-break-after: auto; }");
            sb.AppendLine("</style></head><body><div class='container'>");

            if (CheckIncludeSummary.IsChecked ?? false)
            {
                // --- データ生成：1. サマリセクション ---
                var summaries = RecService.Instance.Evidence.GetSummary();

                sb.AppendLine("<div class='summary-section'>");
                sb.AppendLine("<div class='summary-header'>テスト実施サマリ</div>");
                sb.AppendLine("<table class='summary-table'>");
                sb.AppendLine("<thead><tr><th>ケースNo</th><th>実施日時</th><th>開始時間</th><th>終了時間</th><th>ステップ数</th><th>結果</th></tr></thead><tbody>");

                foreach (var s in summaries)
                {
                    if (s.FinalResult == TestResult.SS) continue;
                    string resClass = s.FinalResult switch { TestResult.OK => "res-ok", TestResult.NG => "res-ng", TestResult.PEND => "res-pend", _ => "" };
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td>No. {s.CaseId}</td>");
                    sb.AppendLine($"<td>{startAt.Add(s.StartTime):yyyy/MM/dd}</td>");
                    sb.AppendLine($"<td>{startAt.Add(s.StartTime):HH:mm:ss}</td>");
                    sb.AppendLine($"<td>{startAt.Add(s.EndTime):HH:mm:ss}</td>");
                    sb.AppendLine($"<td>{s.StepCount} step(s)</td>");
                    sb.AppendLine($"<td class='{resClass}'>{s.FinalResult}</td>");
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</tbody></table></div>");

                // --- データ生成：2. 詳細ログセクション ---
                sb.AppendLine("<h3 style='margin-left: 5px;'>詳細ログ</h3>");
            }

            sb.AppendLine("<div class='table-wrapper'><table>");
            sb.AppendLine("<thead><tr><th class='col-case'>ケース/Step</th><th class='col-note'>コメント</th><th class='col-ss'>スクリーンショット</th></tr></thead><tbody>");

            int lastCaseId = -1;
            int stepCounter = 0;

            foreach (var item in targetItems)
            {
                if (item.OriginalBookmark.CaseId != lastCaseId)
                {
                    lastCaseId = item.OriginalBookmark.CaseId;
                    stepCounter = 1;
                }
                else
                {
                    stepCounter++;
                }

                string timeStr = showRelative
                    ? $"+{(int)item.Time.TotalHours:D2}:{item.Time.Minutes:D2}:{item.Time.Seconds:D2}"
                    : startAt.Add(item.Time).ToString("HH:mm:ss");

                sb.AppendLine("<tr>");
                // ケース/Step情報列
                sb.AppendLine("  <td class='col-case'>");
                sb.AppendLine($"    <span class='case-id'>No {item.OriginalBookmark.CaseId}</span>");
                sb.AppendLine($"    <span class='step-id'>Step {stepCounter}</span>");
                sb.AppendLine($"    <span class='case-time'>{timeStr}</span>");
                sb.AppendLine("  </td>");
                // コメント列
                sb.AppendLine($"  <td class='col-note'>{item.Note}</td>");
                // スクリーンショット列
                var base64String = "data:image/png;base64," + ImageService.BitmapSourceToBase64(item.SnapshotImage);
                sb.AppendLine($"  <td class='col-ss'><img src='{base64String}' class='ss-image'></td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody></table></div></div></body></html>");

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
            bool showRelative = false;
            bool checkIncludeSummary = false;
            this.Dispatcher.Invoke(() =>
            {
                showRelative = RadioRelativeTime.IsChecked ?? true;
                checkIncludeSummary = CheckIncludeSummary.IsChecked ?? true;
            });

            var targetItems = GetTargetItems();
            if (targetItems.Count == 0) return;

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("録画レポート");

            // --- 全体の列幅設定 ---
            ws.Column(1).Width = 15; // ケース/Step用
            ws.Column(2).Width = 25; // コメント用
            ws.Column(3).Width = 90; // スクリーンショット用（大きめに確保）

            var recordDate = RecService.Instance.Evidence.RecordingDate;
            int currentRow = 1;

            if (checkIncludeSummary)
            {
                // --- 1. サマリセクションの出力 ---
                var summaries = RecService.Instance.Evidence.GetSummary();
                var titleRange = ws.Range(currentRow, 1, currentRow, 6);
                titleRange.Merge().Value = "テスト実施サマリ";
                titleRange.Style.Font.Bold = true;
                titleRange.Style.Font.FontSize = 14;
                titleRange.Style.Font.FontColor = XLColor.White;
                titleRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4"); // 濃い青
                currentRow++;

                string[] headers = { "ケースNo", "実施日時", "開始時間", "終了時間", "ステップ数", "結果" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(currentRow, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2"); // 薄いグレー
                    cell.Style.Font.Bold = true;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                currentRow++;

                foreach (var s in summaries)
                {
                    if (s.FinalResult == TestResult.SS) continue;
                    ws.Cell(currentRow, 1).Value = $"No. {s.CaseId}";
                    ws.Cell(currentRow, 2).Value = recordDate.Add(s.StartTime).ToString("yyyy/MM/dd");
                    ws.Cell(currentRow, 3).Value = recordDate.Add(s.StartTime).ToString("HH:mm:ss");
                    ws.Cell(currentRow, 4).Value = recordDate.Add(s.EndTime).ToString("HH:mm:ss");
                    ws.Cell(currentRow, 5).Value = $"{s.StepCount} step(s)";
                    var resCell = ws.Cell(currentRow, 6);
                    resCell.Value = s.FinalResult.ToString();
                    resCell.Style.Font.Bold = true;
                    if (s.FinalResult == TestResult.OK) resCell.Style.Font.FontColor = XLColor.FromHtml("#28A745");
                    else if (s.FinalResult == TestResult.NG) resCell.Style.Font.FontColor = XLColor.FromHtml("#DC3545");
                    ws.Range(currentRow, 1, currentRow, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;
                }

                currentRow += 2; // 余白

                // --- 2. 詳細ログセクションの出力 ---
                ws.Cell(currentRow, 1).Value = "詳細ログ";
                ws.Cell(currentRow, 1).Style.Font.Bold = true;
                ws.Cell(currentRow, 1).Style.Font.FontSize = 12;
                currentRow++;
            }

            int lastCaseId = -1;
            int stepCounter = 0;

            foreach (var item in targetItems)
            {
                if (item.OriginalBookmark.CaseId != lastCaseId)
                {
                    lastCaseId = item.OriginalBookmark.CaseId;
                    stepCounter = 1;
                }
                else
                {
                    stepCounter++;
                }

                // ステップごとの青ヘッダー
                var headerRange = ws.Range(currentRow, 1, currentRow, 3);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4");
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                ws.Cell(currentRow, 1).Value = "ケース/Step";
                ws.Cell(currentRow, 2).Value = "コメント";
                ws.Cell(currentRow, 3).Value = "スクリーンショット";
                currentRow++;

                // データ行の書き込み
                string timeStr = showRelative ? item.Time.ToString(@"hh\:mm\:ss") : recordDate.Add(item.Time).ToString("HH:mm:ss");
                var infoCell = ws.Cell(currentRow, 1);
                infoCell.Value = $"No.{item.OriginalBookmark.CaseId}\nStep {stepCounter}\n{timeStr}";
                infoCell.Style.Alignment.WrapText = true;
                infoCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                var noteCell = ws.Cell(currentRow, 2);
                noteCell.Value = item.Note;
                noteCell.Style.Alignment.WrapText = true;
                noteCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                // --- 画像挿入と行高さ調整 ---
                if (item.SnapshotImage != null)
                {
                    var imgData = ImageData.GetImageData(item.SnapshotImage);
                    if (imgData != null && imgData.Bytes != null && imgData.Bytes.Length > 0)
                    {
                        using (var imageStream = new MemoryStream(imgData.Bytes))
                        {
                            imageStream.Seek(0, SeekOrigin.Begin);
                            var picture = ws.AddPicture(imageStream);

                            // 1. 位置を確定
                            picture.MoveTo(ws.Cell(currentRow, 3), 5, 5);
                            picture.Placement = XLPicturePlacement.Move;

                            // 2. サイズの設定（ScaleやPixelHeightが0の場合の回避策）
                            double dpiScale = VisualTreeHelper.GetDpi(this).DpiScaleX;

                            // Scaleが0なら1.0(等倍)として扱う
                            double safeScale = item.Scale > 0 ? item.Scale : 1.0;

                            int targetWidth = (int)Math.Round((imgData.PixelWidth * safeScale) / dpiScale);
                            int targetHeight = (int)Math.Round((imgData.PixelHeight * safeScale) / dpiScale);

                            // もし計算結果が極端に小さい場合は、デフォルトサイズを適用
                            if (targetHeight < 50)
                            {
                                targetWidth = 400; // 仮の幅
                                targetHeight = 225; // 仮の高さ（16:9）
                            }

                            picture.Width = targetWidth;
                            picture.Height = targetHeight;

                            // 3. ★最重要★ 行の高さを強制確保
                            // ポイント単位への変換(0.75)に余裕(+20)を持たせる
                            double rowHeightPt = (targetHeight * 0.75) + 20;

                            // 最低でも150ポイント以上の高さを確保するようにして、画像を見えるようにする
                            ws.Row(currentRow).Height = Math.Max(rowHeightPt, 150);

                            //Debug.WriteLine($"画像配置完了: Row={currentRow}, CalcHeight={targetHeight}, RowHeightSet={ws.Row(currentRow).Height}");
                        }
                    }
                }
                else
                {
                    ws.Row(currentRow).Height = 40; // 画像がない場合の最小高さ
                }

                currentRow += 2; // 次のステップとの間に空行を入れる
            }

            workbook.SaveAs(fullPath);
        }

        private async void StartCapture_Click(object sender, RoutedEventArgs e)
        {
            if(Owner is MainWindow main)
            {
                var scale = GetSelectedScale();
                var items = ExportItems.Where(x => x.IsSelected);
                int index = 0;
                await RunExportTask(async (progress) =>
                {
                    foreach (var item in items)
                    {
                        var cached = _cacheManager.GetCachedImage(item.OriginalBookmark.Id, scale);
                        if (cached is null || item.OriginalBookmark.IsDirty)
                        {
                            main.VideoPlayer.Position = item.OriginalBookmark.Time;
                            await Task.Delay(500);
                            var snapshot = new VideoSnapshotInfo(main.VideoPlayer);
                            var result = RecService.Instance.SaveImage(item.OriginalBookmark, snapshot, scale);
                            if (result != null && result.Value.Bitmap != null)
                            {
                                _cacheManager.RegisterCache(item.OriginalBookmark.Id, result.Value.Bitmap, scale);
                                item.SnapshotImage = result.Value.Bitmap;
                                item.OriginalBookmark.IsDirty = false;
                            }
                        }
                        else
                        {
                            item.SnapshotImage = cached;
                        }
                        progress.Report((++index) * 100 / items.Count());
                    }
                });
                StatusText.Text = "撮影完了。出力する画像を選択・確認してください。";
                OutputGroupBox.IsEnabled = true;
            }
        }

        private void RefreshExportIds()
        {
            int currentCaseId = -1;
            int currentStepId = 0;

            // 現在の並び順（Order -> Time）で、チェックが入っているものだけを対象にする
            var activeItems = ExportItems
                .Where(x => x.IsSelected)
                .OrderBy(x => x.Order)
                .ThenBy(x => x.Time);

            foreach (var item in activeItems)
            {
                // CaseIdが変わったらStepIdを1リセット、同じならインクリメント
                currentStepId = (currentCaseId == item.OriginalBookmark.CaseId) ? currentStepId + 1 : 1;
                currentCaseId = item.OriginalBookmark.CaseId;

                item.CaseId = currentCaseId;
                item.StepId = currentStepId;
                item.Note = string.IsNullOrEmpty(item.OriginalBookmark.Note) ? " " 
                    : item.OriginalBookmark.Note.Replace("\r\n", " ").Replace("\n", " ");
            }
        }

        private List<ExportItemViewModel> GetTargetItems()
        {
            return ExportItems.Where(x => x.IsSelected).ToList();
        }

        private void Item_MouseMove(object sender, MouseEventArgs e)
        {
            // 左クリックが押されている時だけドラッグ開始
            if (e.LeftButton == MouseButtonState.Pressed && sender is Border border)
            {
                var draggedItem = border.DataContext;
                if (draggedItem == null) return;

                // データオブジェクトを作成
                var dataObject = new DataObject("TiledImage", draggedItem);

                // 【重要】ここでドラッグ開始。このメソッドが終わるまで処理は止まります。
                DragDrop.DoDragDrop(border, dataObject, DragDropEffects.Move);
            }
        }

        private void Thumbnail_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // ダブルクリック（2回連続クリック）の時だけ実行
            if (e.ClickCount == 2)
            {
                if (sender is FrameworkElement elem && elem.DataContext is ExportItemViewModel item)
                {
                    _currentIdx = ExportPreviewList.Items.IndexOf(item);

                    // プレビューモードへ移行
                    _isPreviewMode = true;
                    EnterPreviewMode();

                    // イベントを完了させる（ドラッグ処理などへ流さない）
                    e.Handled = true;
                }
            }
        }

        private void EnterPreviewMode()
        {
            _isPreviewMode = true;

            ExportPreviewList.Visibility = Visibility.Collapsed;
            PreviewArea.Visibility = Visibility.Visible;

            UpdatePreviewDisplay();
            // UIの表示切り替え（Bindingを使っている場合はProperty通知）
        }

        private void UpdatePreviewDisplay()
        {
            var allItems = ExportPreviewList.Items.Cast<ExportItemViewModel>().ToList();
            var selectedItems = allItems.Where(x => x.IsSelected).ToList();

            if (_currentIdx >= 0 && _currentIdx < allItems.Count)
            {
                var item = allItems[_currentIdx];
                // 1. 重要：XAMLのバインド先であるプロパティにアイテムをセットする
                // これにより、追加したTextBlock等の値が自動的に更新されます
                SelectedPreviewItem = item;

                // 2. カウンターのテキスト更新（ViewModelのプロパティを更新するか、直接書き換える）
                int displayIdx = selectedItems.IndexOf(item) + 1;
                string counterText = $"{displayIdx} / {selectedItems.Count}";

                // ViewModel側にプロパティがあるならそちらを更新
                PreviewCounterText = counterText;
                // もし直書きなら
                //PreviewPageCounter.Text = counterText;

                // 3. 画像の更新（バインドが正しく動いていれば不要ですが、念のため）
                LargePreviewImage.Source = item?.SnapshotImage;
            }
        }

        private void ClosePreview_Click(object sender, MouseButtonEventArgs e)
        {
            _isPreviewMode = false;
            PreviewArea.Visibility = Visibility.Collapsed;
            ExportPreviewList.Visibility = Visibility.Visible;
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (!_isPreviewMode)
            {
                base.OnKeyDown(e);
                return;
            }

            var allItems = ExportPreviewList.Items.Cast<ExportItemViewModel>().ToList();
            if (allItems.Count == 0) return;

            var currentItem = allItems[_currentIdx];

            switch (e.Key)
            {
                case Key.Left:
                    int prevIdx = allItems.Take(_currentIdx).ToList().FindLastIndex(x => x.IsSelected);
                    if (prevIdx != -1) MoveTo(prevIdx, false, false);
                    e.Handled = true;
                    break;

                case Key.Right:
                    int nextIdxRel = allItems.Skip(_currentIdx + 1).ToList().FindIndex(x => x.IsSelected);
                    if (nextIdxRel != -1) MoveTo(_currentIdx + 1 + nextIdxRel, true, false);
                    e.Handled = true;
                    break;

                case Key.Up:
                    int upIdx = allItems.Take(_currentIdx).ToList()
                                        .FindLastIndex(x => x.IsSelected && x.OriginalBookmark.CaseId != currentItem.OriginalBookmark.CaseId);
                    if (upIdx != -1)
                    {
                        var targetCaseId = allItems[upIdx].OriginalBookmark.CaseId;
                        int firstInCase = allItems.FindIndex(x => x.IsSelected && x.OriginalBookmark.CaseId == targetCaseId);
                        if (firstInCase != -1) MoveTo(firstInCase, false, true);
                    }
                    e.Handled = true;
                    break;

                case Key.Down:
                    int downIdxRel = allItems.Skip(_currentIdx + 1).ToList()
                                             .FindIndex(x => x.IsSelected && x.OriginalBookmark.CaseId != currentItem.OriginalBookmark.CaseId);
                    if (downIdxRel != -1) MoveTo(_currentIdx + 1 + downIdxRel, true, true);
                    e.Handled = true;
                    break;

                case Key.Escape:
                    _isPreviewMode = false;
                    PreviewArea.Visibility = Visibility.Collapsed;
                    ExportPreviewList.Visibility = Visibility.Visible;
                    e.Handled = true;
                    break;
            }

            if (!e.Handled)
            {
                base.OnPreviewKeyDown(e);
            }
        }

        private void MoveTo(int newIdx, bool isForward, bool isVertical)
        {
            _currentIdx = newIdx;
            UpdatePreviewDisplay();
            AnimateSlide(isForward, isVertical);
        }

        private void AnimateSlide(bool isForward, bool isVertical)
        {
            // 方向に応じた移動距離を設定 (左右ならX, 上下ならY)
            double distance = 100; // スライド量
            double startValue = isForward ? distance : -distance;

            var animation = new DoubleAnimation
            {
                From = startValue,
                To = 0,
                Duration = TimeSpan.FromMilliseconds(250),
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };

            if (isVertical)
            {
                // 上下アニメーション (Y軸)
                SlideTransform.X = 0; // Xをリセット
                SlideTransform.BeginAnimation(TranslateTransform.YProperty, animation);
            }
            else
            {
                // 左右アニメーション (X軸)
                SlideTransform.Y = 0; // Yをリセット
                SlideTransform.BeginAnimation(TranslateTransform.XProperty, animation);
            }
        }
    }
}
