namespace TraceShot.ViewModels
{
    using CommunityToolkit.Mvvm.ComponentModel;
    using System.Windows.Media.Imaging;
    using TraceShot.Models;

    public partial class ExportItemViewModel : ObservableObject
    {
        // 参照用
        public Bookmark OriginalBookmark { get; set; }

        // プレビュー用画像（撮影直後にメモリ上のBitmapをセット）
        public BitmapSource SnapshotImage { get; set; }

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
