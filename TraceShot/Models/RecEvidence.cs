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
    private ObservableCollection<TimelineEntry> _bookmarks = [];

    partial void OnIsCropEnabledChanged(bool value)
    {
        foreach (var b in Bookmarks) b.MarkAsDirty();
    }
}