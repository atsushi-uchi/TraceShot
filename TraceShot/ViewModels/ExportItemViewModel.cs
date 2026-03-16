namespace TraceShot.ViewModels
{
    using CommunityToolkit.Mvvm.ComponentModel;
    using System.Windows;
    using System.Windows.Media.Imaging;
    using TraceShot.Models;
    using TraceShot.Services;

    public partial class ExportItemViewModel : ObservableObject
    {
        public string EditableNote
        {
            get => OriginalBookmark.Note;
            set
            {
                if (OriginalBookmark.Note != value)
                {
                    OriginalBookmark.Note = value;
                    OriginalBookmark.IsDirty = true; // 編集されたら再出力フラグを立てる
                    OnPropertyChanged();
                }
            }
        }

        public ExportItemViewModel(Bookmark bookmark, BitmapSource image)
        {
            this.OriginalBookmark = bookmark;
            this.SnapshotImage = image;

            // Bookmark 側の情報を反映
            // もし Note が変更されたら Bookmark 側も更新したい場合は、
            // ここで双方向の工夫が必要になりますが、まずはシンプルに同期します。
        }

        // デザイン時や、空のアイテムが必要な場合用のデフォルトコンストラクタ
        public ExportItemViewModel() { }

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
    }
}
