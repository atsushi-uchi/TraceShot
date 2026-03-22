namespace TraceShot.ViewModels
{
    using CommunityToolkit.Mvvm.ComponentModel;
    using System.Windows.Media.Imaging;
    using TraceShot.Models;
    using TraceShot.Services;

    public partial class ExportItemViewModel : ObservableObject
    {
        private readonly ExportCacheManager _cacheManager;
        public ExportItemViewModel(TimelineEntry bookmark, BitmapSource? image, ExportCacheManager cacheManager)
        {
            OriginalBookmark = bookmark;
            SnapshotImage = image;
            _cacheManager = cacheManager;

            _isSelected = _cacheManager.IsPreviouslySelected(bookmark.Id);
        }

        partial void OnIsSelectedChanged(bool value)
        {
            // チェックが変わるたびにマネージャーに記録
            OriginalBookmark.IsExportEnabled = value;
            RecService.Instance.SaveEvidenceJson();
        }

        //public string EditableNote
        //{
        //    get => OriginalBookmark.Note;
        //    set
        //    {
        //        if (OriginalBookmark.Note != value)
        //        {
        //            OriginalBookmark.Note = value;
        //            OriginalBookmark.IsDirty = true;
        //            OnPropertyChanged();
        //            OnPropertyChanged(nameof(Note));
        //        }
        //    }
        //}

        // 参照用
        public TimelineEntry OriginalBookmark { get; set; }

        public string ImagePath { get; set; }

        [ObservableProperty] private BitmapSource _snapshotImage;

        [ObservableProperty] private bool _isSelected = true;

        [ObservableProperty] private int _order;

        [ObservableProperty] private double _scale;

        [ObservableProperty] private string _serialNumber = "";

        [ObservableProperty] private int _caseId;
        [ObservableProperty] private int _stepId;

        [ObservableProperty] private string _note;

        public TimeSpan Time => OriginalBookmark.Time;

        //public string Note => OriginalBookmark.Note ?? "";
        public ExportItemViewModel() { }
    }
}
