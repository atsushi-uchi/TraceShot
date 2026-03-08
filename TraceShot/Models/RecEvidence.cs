using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
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
    private bool _isCropLocked = true;

    // 💡 リスト自体を ObservableProperty にする
    [ObservableProperty]
    private ObservableCollection<Bookmark> _bookmarks = [];
}