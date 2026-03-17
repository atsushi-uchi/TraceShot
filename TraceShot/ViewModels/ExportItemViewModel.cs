namespace TraceShot.ViewModels
{
    using CommunityToolkit.Mvvm.ComponentModel;
    using System.Windows.Media.Imaging;
    using TraceShot.Models;
    using TraceShot.Services;

    public partial class ExportItemViewModel : ObservableObject
    {
        private readonly ExportCacheManager _cacheManager;
        public ExportItemViewModel(Bookmark bookmark, BitmapSource? image, ExportCacheManager cacheManager)
        {
            OriginalBookmark = bookmark;
            SnapshotImage = image;
            _cacheManager = cacheManager;

            _isSelected = _cacheManager.IsPreviouslySelected(bookmark.Id);
        }

        partial void OnIsSelectedChanged(bool value)
        {
            // チェックが変わるたびにマネージャーに記録
            _cacheManager?.UpdateSelection(OriginalBookmark.Id, value);
        }

        public string EditableNote
        {
            get => OriginalBookmark.Note;
            set
            {
                if (OriginalBookmark.Note != value)
                {
                    OriginalBookmark.Note = value;
                    OriginalBookmark.IsDirty = true;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(Note));
                }
            }
        }

        // 参照用
        public Bookmark OriginalBookmark { get; set; }

        // プレビュー用画像（撮影直後にメモリ上のBitmapをセット）
        [ObservableProperty]
        private BitmapSource _snapshotImage;

        // ファイルパス
        public string ImagePath { get; set; }

        // 出力対象にするかどうか（UIのチェックボックスと連動）
        [ObservableProperty]
        private bool _isSelected = true;

        // 並び順（後でドラッグ＆ドロップ実装時に使用）
        [ObservableProperty]
        private int _order;

        [ObservableProperty]
        private double _scale;

        public TimeSpan Time => OriginalBookmark.Time;

        public string Note => OriginalBookmark.Note ?? "";
        public ExportItemViewModel() { }
    }
}
