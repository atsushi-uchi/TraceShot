using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using TraceShot.Controls;
using TraceShot.Extensions;
using TraceShot.Models;
using TraceShot.Services;
using Clipboard = System.Windows.Clipboard;
using DataObject = System.Windows.DataObject;
using IDataObject = System.Windows.IDataObject;

namespace TraceShot.ViewModels
{
    public enum AppViewMode
    {
        Recording, // 録画モード
        Edit,      // 通常編集モード（動画あり）
        Rescue     // 非常事態モード（静止画のみ）
    }

    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty] private AnnotationManager _annotationManager;
        private const string ClipboardFormat = "TraceShot_Annotation_JSON";

        public MainViewModel()
        {
            SetupTimelineView();
            
            _annotationManager = new AnnotationManager();
            _annotationManager.RequestBookmarkSelection += OnRequestBookmarkSelection;

            RecService.Instance.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(RecService.Evidence))
                {
                    App.Current.Dispatcher.Invoke(() =>
                    {
                        SetupTimelineView();
                        UpdateTimelineGroups();
                    });
                }
                if (e.PropertyName == nameof(RecService.IsRecording))
                {
                    OnPropertyChanged(nameof(CanAddEntry));
                }
            };
        }

        private void OnRequestBookmarkSelection(Guid bookmarkId)
        {
            // 現在のリストから該当するブックマークを探して選択
            var target = RecService.Instance.GetBookmark(bookmarkId);
            if (target != null)
            {
                SelectedItem = target;
            }
        }

        public void SetupTimelineView()
        {
            TimelineView = CollectionViewSource.GetDefaultView(TimelineEntries);
            if (TimelineView == null) return;

            // グループ化の設定
            TimelineView.GroupDescriptions.Clear();
            TimelineView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(Bookmark.GroupName)));

            // ソートの設定
            TimelineView.SortDescriptions.Clear();
            TimelineView.SortDescriptions.Add(new SortDescription(nameof(Bookmark.Time), ListSortDirection.Ascending));

            // ライブシェイピングの設定（編集時に即座に並び替える）
            if (TimelineView is ICollectionViewLiveShaping liveView)
            {
                liveView.IsLiveSorting = true;
                liveView.LiveSortingProperties.Clear();
                liveView.LiveSortingProperties.Add(nameof(Bookmark.Time));

                // ケースを跨ぐ移動（IsCaseStartの変更）も即座に反映したい場合
                liveView.IsLiveGrouping = true;
                liveView.LiveGroupingProperties.Clear();
                liveView.LiveGroupingProperties.Add(nameof(Bookmark.CaseId));
            }
        }


        public SettingsService Config => SettingsService.Instance;

        public RecService Recorder => RecService.Instance;

        public ObservableCollection<Bookmark> TimelineEntries => RecService.Instance.Entries;

        [ObservableProperty]  private ICollectionView? _timelineView;

        [ObservableProperty] private Bookmark? _selectedItem;

        [ObservableProperty] private AppViewMode _currentMode = AppViewMode.Recording;

        public bool IsEditMode
        {
            get => CurrentMode == AppViewMode.Edit || CurrentMode == AppViewMode.Rescue;
            set
            {
                if (value)
                {
                    //CurrentMode = AppViewMode.Edit;

                    // 物理ファイルの存在チェックで Edit か Rescue かを最終決定
                    string videoPath = Path.Combine(RecService.Instance.CurrentFolder, RecService.Instance.Evidence?.VideoFileName ?? "");

                    if (File.Exists(videoPath))
                        CurrentMode = AppViewMode.Edit;
                    else
                        CurrentMode = AppViewMode.Rescue;
                }
                else
                {
                    // OFFにした時：録画モードへ
                    CurrentMode = AppViewMode.Recording;
                }
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentMode));
            }
        }

        [ObservableProperty] private int _nextNo = 1;

        public Action<Bookmark>? ScrollIntoViewRequested { get; set; }
        public Action? RefreshCanvas { get; set; }
        public Func<TimeSpan>? GetCurrentPosition { get; set; }
        public Func<VideoSnapshotInfo>? GetVideoSnapshotFunc { get; set; }

        public event EventHandler? ShutterRequested;

        [ObservableProperty] private Uri? _videoSource;
        [ObservableProperty] private string _statusText = "";
        [ObservableProperty] private bool _canAddEntry = false;
        [ObservableProperty] private BitmapSource? _rescueImageSource;

        public async Task<bool> LoadEvidenceAsync(string filePath)
        {
            try
            {
                string jsonString = File.ReadAllText(filePath);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var loaded = JsonSerializer.Deserialize<RecEvidence>(jsonString, options);
                if (loaded == null) throw new Exception("JSONファイルのデシリアライズに失敗");

                foreach (var bookmark in loaded.Entries)
                {
                    foreach (var annotation in bookmark.Annotations)
                    {
                        annotation.IsSelected = false;

                        if (annotation is RectAnnotation rect)
                        {
                            rect.OcrAction = ExecuteOcrAction;
                            rect.PropertyChanged += (s, e) =>
                            {
                                if (e.PropertyName == nameof(RectAnnotation.IsFocused))
                                {
                                    AnnotationManager.RefreshCropOverlay();
                                }
                            };
                        }
                    }
                }
                RecService.Instance.Evidence = loaded;
                RecService.Instance.JsonPath = filePath;

                var folderPath = Path.GetDirectoryName(filePath) ?? "";
                RecService.Instance.CurrentFolder = folderPath;
                string videoPath = Path.Combine(folderPath, loaded.VideoFileName ?? "");

                if (File.Exists(videoPath))
                {
                    VideoSource = new Uri(videoPath);
                    StatusText = $"読み込み: {loaded.RecMode} {loaded.VideoFileName}";
                }
                else
                {
                    StatusText = "動画が見つかりません。静止画編集モードで起動します。";
                    var firstEntry = loaded.Entries.OrderBy(x => x.Time).FirstOrDefault();
                    if (firstEntry != null && File.Exists(firstEntry.ImagePath))
                    {
                        var bitmap = LoadImageFromFile(firstEntry.ImagePath);
                        if (bitmap != null)
                        {
                            RescueImageSource = bitmap;
                        }
                    }
                }
                IsEditMode = true;
                UpdateTimelineGroups();
                CanAddEntry = true;
                return true;
            }
            catch (Exception ex)
            {
                StatusText = $"ファイル読込失敗 {ex.Message}";
                return false;
            }
        }

        public BitmapSource? LoadImageFromFile(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return null;

            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad; // ファイルロックを回避
                bitmap.UriSource = new Uri(path, UriKind.Absolute);
                bitmap.EndInit();
                bitmap.Freeze(); // スレッド間での共有を可能にする
                return bitmap;
            }
            catch
            {
                return null;
            }
        }

        public async Task ExecuteFocusAction(RectAnnotation rect)
        {
            AnnotationManager.RefreshCropOverlay();
        }

        public async Task ExecuteOcrAction(RectAnnotation rect)
        {
            if (SelectedItem == null) return;

            // 1. ビデオ情報を取得
            //var info = new VideoSnapshotInfo(VideoPlayer);
            var info = GetVideoSnapshotFunc?.Invoke();
            if (info == null) return;

            // 2. 正味の画像を生成（既存の ImageService を利用）
            BitmapSource pureVideoBitmap = ImageService.GeneratePureVideoBitmap(SelectedItem, info);

            // 3. RectAnnotation の相対座標（RelX, RelY...）を使用
            // 0.0〜1.0 なので、そのまま PixelWidth/Height を掛けるだけ
            int px = (int)(rect.RelX * pureVideoBitmap.PixelWidth);
            int py = (int)(rect.RelY * pureVideoBitmap.PixelHeight);
            int pw = (int)(rect.RelWidth * pureVideoBitmap.PixelWidth);
            int ph = (int)(rect.RelHeight * pureVideoBitmap.PixelHeight);

            // 範囲チェック（画像の外にはみ出さないように）
            px = Math.Clamp(px, 0, pureVideoBitmap.PixelWidth - 1);
            py = Math.Clamp(py, 0, pureVideoBitmap.PixelHeight - 1);
            pw = Math.Min(pw, pureVideoBitmap.PixelWidth - px);
            ph = Math.Min(ph, pureVideoBitmap.PixelHeight - py);

            if (pw <= 0 || ph <= 0) return;

            var cropped = new CroppedBitmap(pureVideoBitmap, new Int32Rect(px, py, pw, ph));

            // 4. OCR実行
            string result = await ImageService.RecognizeTextFromBitmapSource(cropped);

            if (!string.IsNullOrWhiteSpace(result))
            {
                string cleanText = result.Replace("\r", "").Replace("\n", " ").Trim();
                SelectedItem.AddNewLine(cleanText);
            }
        }

        public void AddTimelineEntry()
        {
            ExecuteResult(TestResult.SS.ToString());
        }

        public void AddVoiceEntry()
        {
            ExecuteResult(TestResult.SS.ToString());
        }

        [RelayCommand]
        private void ExecuteResult(string result)
        {
            if (!Enum.TryParse<TestResult>(result, out var resultType)) return;

            switch (CurrentMode)
            {
                case AppViewMode.Recording:
                    RecExecuteResult(resultType);
                    break;
                case AppViewMode.Edit:
                    EditExecuteResult(resultType);
                    break;
                case AppViewMode.Rescue:
                    RescueExecuteResult(resultType);
                    break;
            }
        }

        private void RescueExecuteResult(TestResult resultType)
        {
            if (SelectedItem is Bookmark entry)
            {
                entry.Result = resultType;
                UpdateTimelineGroups();
                RefreshCanvas?.Invoke();
            }
        }

        private void EditExecuteResult(TestResult resultType)
        {
            if (SelectedItem is Bookmark entry)
            {
                entry.Result = resultType;
                UpdateTimelineGroups();
                RefreshCanvas?.Invoke();
            }
            else
            {
                var currentTime = GetCurrentPosition?.Invoke();
                if (currentTime.HasValue)
                {
                    int caseId = 0;
                    var lastEntry = TimelineEntries.FirstOrDefault(x => x.Time > currentTime);
                    if (lastEntry != null)
                    {
                        caseId = lastEntry.CaseId;
                    }

                    entry = new()
                    {
                        Time = currentTime.Value,
                        CaseId = caseId,
                        Result = resultType,
                        Note = "",
                        Icon = "📸"
                    };
                    ShutterRequested?.Invoke(this, EventArgs.Empty);
                    TimelineEntries.Add(entry);
                    UpdateTimelineGroups();
                    RefreshCanvas?.Invoke();
                    ScrollIntoViewRequested?.Invoke(entry);
                }
            }
        }

        private void RecExecuteResult(TestResult resultType)
        {
            var lastEntry = TimelineEntries.LastOrDefault();
            var noteText = resultType.In(TestResult.SS) ? "" : $"No.{NextNo} {resultType}";
            var isCaseStart = lastEntry?.Result.In(TestResult.OK, TestResult.NG, TestResult.PEND) ?? true;

            Bookmark entry = new()
            {
                Time = Recorder.CurrentDuration,
                CaseId = NextNo,
                Result = resultType,
                Note = noteText,
                Icon = "📸"
            };
            Recorder.SaveBitmap(entry);

            if (resultType.In(TestResult.OK, TestResult.NG, TestResult.PEND))
            {
                NextNo++;
            }

            ShutterRequested?.Invoke(this, EventArgs.Empty);

            TimelineEntries.Add(entry);
            UpdateTimelineGroups();
            ScrollIntoViewRequested?.Invoke(entry);
        }

        public void UpdateTimelineGroups()
        {
            var sortedList = TimelineEntries.OrderBy(x => x.Time).ToList();
            Dictionary<int, int> caseIds = [];
            string groupName = "";
            int prevId = int.MinValue;

            foreach (var entry in sortedList)
            {
                if (prevId != entry.CaseId)
                {
                    if (caseIds.ContainsKey(entry.CaseId))
                    {
                        int count = ++caseIds[entry.CaseId];
                        groupName = (entry.CaseId == 0) ? $"--_{count}" : $"No.{entry.CaseId}_{count}";
                    }
                    else
                    {
                        caseIds[entry.CaseId] = 0;
                        groupName = (entry.CaseId == 0) ? "--" : $"No.{entry.CaseId}";
                    }
                    prevId = entry.CaseId;
                }
                entry.GroupName = groupName;
            }
            App.Current.Dispatcher.Invoke(() =>
            {
                TimelineView?.Refresh();
                //if (SelectedItem != null)
                //{
                //    ScrollIntoViewRequested?.Invoke(SelectedItem);
                //}
            });
        }

        public void CopyAnnotation(AnnotationBase? target = null)
        {
            var targetAnnotation = target ?? AnnotationManager.SelectedAnnotation;
            if (targetAnnotation is AnnotationBase finalTarget)
            {
                try
                {
                    string json = JsonSerializer.Serialize<AnnotationBase>(finalTarget);

                    var data = new DataObject();
                    data.SetData("TraceShot_Annotation_JSON", json);
                    Clipboard.SetDataObject(data);

                    StatusText = "注釈をクリップボードにコピーしました";
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Copy failed: {ex.Message}");
                }
            }
        }

        public void PasteAnnotation(Bookmark bookmark)
        {
            IDataObject data = Clipboard.GetDataObject();
            if (data.GetDataPresent(ClipboardFormat))
            {
                if (data.GetData(ClipboardFormat) is not string json) return;
                var paseted = DeserializeAnnotation(json);
                if (paseted != null)
                {
                    AnnotationManager.AddPastedAnnotation(bookmark, paseted);
                    bookmark.Modified();
                }
            }
        }

        private AnnotationBase? DeserializeAnnotation(string json)
        {
            try
            {
                return JsonSerializer.Deserialize<AnnotationBase>(json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Deserialization failed: {ex.Message}");
                return null;
            }
        }

        public void SwitchToRescueMode(string error)
        {
            var bookmark = TimelineEntries.OrderBy(x => x.Time).FirstOrDefault();
            if (File.Exists(bookmark?.ImagePath))
            {
                var bitmap = LoadImageFromFile(bookmark.ImagePath);
                RescueImageSource = bitmap;
                CurrentMode = AppViewMode.Rescue;
                StatusText = $"録画に失敗したため、救済モードに切り替えます。エラー内容：{error}";
            }
            else
            {
                StatusText = $"録画に失敗しました。 エラー内容：{error}";
            }
        }
    }
}
