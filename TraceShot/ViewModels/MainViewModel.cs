using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using TraceShot.Extensions;
using TraceShot.Models;
using TraceShot.Services;

namespace TraceShot.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        public MainViewModel()
        {
            SetupTimelineView();

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
            };
        }

        public void SetupTimelineView()
        {
            TimelineView = CollectionViewSource.GetDefaultView(TimelineEntries);
            if (TimelineView == null) return;

            // グループ化の設定
            TimelineView.GroupDescriptions.Clear();
            TimelineView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(TimelineEntry.GrpupName)));

            // ソートの設定
            TimelineView.SortDescriptions.Clear();
            TimelineView.SortDescriptions.Add(new SortDescription(nameof(TimelineEntry.Time), ListSortDirection.Ascending));

            // ライブシェイピングの設定（編集時に即座に並び替える）
            if (TimelineView is ICollectionViewLiveShaping liveView)
            {
                liveView.IsLiveSorting = true;
                liveView.LiveSortingProperties.Clear();
                liveView.LiveSortingProperties.Add(nameof(TimelineEntry.Time));

                // ケースを跨ぐ移動（IsCaseStartの変更）も即座に反映したい場合
                liveView.IsLiveGrouping = true;
                liveView.LiveGroupingProperties.Clear();
                liveView.LiveGroupingProperties.Add(nameof(TimelineEntry.CaseId));
            }
        }


        public SettingsService Config => SettingsService.Instance;

        public RecService Recorder => RecService.Instance;

        public ObservableCollection<TimelineEntry> TimelineEntries => RecService.Instance.Entries;
        [ObservableProperty]  private ICollectionView? _timelineView;


        public Action<TimelineEntry>? ScrollIntoViewRequested { get; set; }
        public Action? RefreshDisplay { get; set; }
        public Func<TimeSpan>? GetCurrentPosition { get; set; }



        [ObservableProperty] private TimelineEntry? _selectedItem;


        [ObservableProperty] private bool _isEditMode = false;

        [ObservableProperty] private int _nextNo = 1;

        public WriteableBitmap? PreviewBitmap;

        [RelayCommand]
        private void ExecuteResult(string result)
        {
            if (!Enum.TryParse<TestResult>(result, out var resultType)) return;

            if (IsEditMode)
            {
                EditExecuteResult(resultType);
            }
            else
            {
                RecExecuteResult(resultType);
            }
        }

        private void EditExecuteResult(TestResult resultType)
        {
            if (SelectedItem is TimelineEntry entry)
            {
                entry.Result = resultType;
                UpdateTimelineGroups();
                RefreshDisplay?.Invoke();
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
                    SoundService.Instance.PlayShutter();
                    TimelineEntries.Add(entry);
                    UpdateTimelineGroups();
                    RefreshDisplay?.Invoke();
                }
            }
        }


        private void RecExecuteResult(TestResult resultType)
        {
            var lastEntry = TimelineEntries.LastOrDefault();
            var noteText = resultType.In(TestResult.SS) ? "" : $"No.{NextNo} {resultType}";
            var isCaseStart = lastEntry?.Result.In(TestResult.OK, TestResult.NG, TestResult.PEND) ?? true;

            TimelineEntry entry = new()
            {
                Time = Recorder.CurrentDuration,
                CaseId = NextNo,
                Result = resultType,
                Note = noteText,
                Icon = "📸"
            };

            if (resultType.In(TestResult.OK, TestResult.NG, TestResult.PEND))
            {
                NextNo++;
            }

            SoundService.Instance.PlayShutter();

            if (PreviewBitmap != null)
            {
                var path = Recorder.SaveBitmap(entry, PreviewBitmap);
                entry.ImagePath = path;
            }

            TimelineEntries.Add(entry);
            UpdateTimelineGroups();
            SelectedItem = entry;
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
                        groupName = (entry.CaseId == 0) ? $"Screen Shot_{count}" : $"No.{entry.CaseId}_{count}";
                    }
                    else
                    {
                        caseIds[entry.CaseId] = 0;
                        groupName = (entry.CaseId == 0) ? "Screen Shot" : $"No.{entry.CaseId}";
                    }
                    prevId = entry.CaseId;
                }
                entry.GrpupName = groupName;
            }
            App.Current.Dispatcher.Invoke(() =>
            {
                TimelineView.Refresh();
            });
        }
    }
}
