using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Windows;
using TraceShot.Models;

public partial class RecEvidence : ObservableObject
{
    [ObservableProperty]
    private DateTime _recordingDate;

    [ObservableProperty]
    private string _videoFileName = "";

    [ObservableProperty]
    private string _recMode = "";

    [ObservableProperty]
    private CropState _cropState = CropState.Editing;

    [ObservableProperty]
    private bool _isCropEnabled = false;

    [ObservableProperty]
    private Rect _commonCropRect = new Rect(0.1, 0.1, 0.8, 0.8);

    [ObservableProperty]
    private ObservableCollection<TimelineEntry> _entries = [];

    partial void OnIsCropEnabledChanged(bool value)
    {
        foreach (var b in Entries) b.EntryAsDirty();
    }

    public List<CaseSummary> GetSummary()
    {
        return Entries
            .GroupBy(e => e.CaseId)
            .Select(g =>
            {
                var sortedInCase = g
                    .Where(e => e.IsExportEnabled)
                    .OrderBy(e => e.Time).ToList();

                // 要素がない場合は空のリストとして扱う（nullを返さない）
                if (sortedInCase.Count == 0) return Enumerable.Empty<CaseSummary>();

                var firstEntry = sortedInCase.First();
                var lastEntry = sortedInCase.Last();

                var summary = new CaseSummary
                {
                    CaseId = g.Key,
                    StepCount = sortedInCase.Count,
                    FinalResult = lastEntry.Result,
                    StartTime = firstEntry.Time,
                    EndTime = lastEntry.Time,
                };

                return new[] { summary }; // 1要素の配列として返す
            })
            .SelectMany(s => s) // ここでフラットに展開（空なら自動で除外される）
            .OrderBy(s => s.StartTime)
            .ToList();
    }
}