using CommunityToolkit.Mvvm.ComponentModel;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    public partial class NoteAnnotation : AnnotationBase
    {
        [ObservableProperty] private double _startX;
        [ObservableProperty] private double _startY;
        [ObservableProperty] private double _relStartX;
        [ObservableProperty] private double _relStartY;
        [ObservableProperty] private string _text = string.Empty;
        [ObservableProperty] private bool _isEditing;
        [ObservableProperty] private bool _isCommitted;
        [ObservableProperty] private double _actualTextWidth = 120; // 初期値（目安）
        [ObservableProperty] private double _actualTextHeight = 60; // 初期値（目安）

        public string OriginText = string.Empty;
        partial void OnRelStartXChanged(double value) => StartX = value * _lastParentWidth;
        partial void OnRelStartYChanged(double value) => StartY = value * _lastParentHeight;

        public NoteAnnotation()
        {
            this.PropertyChanged += (s, e) =>
            {
                // 親のサイズ(size)をどこかに保持している前提
                if (e.PropertyName == nameof(RelX)) X = RelX * _lastParentWidth;
                if (e.PropertyName == nameof(RelY)) Y = RelY * _lastParentHeight;
                if (e.PropertyName == nameof(RelStartX)) StartX = RelStartX * _lastParentWidth;
                if (e.PropertyName == nameof(RelStartY)) StartY = RelStartY * _lastParentHeight;
            };
        }

        public override void OnStart(Point pos, Size size)
        {
            _lastParentWidth = size.Width;
            _lastParentHeight = size.Height;
            base.OnStart(pos, size);

            RelStartX = Normalize(pos.X, size.Width);
            RelStartY = Normalize(pos.Y, size.Height);

            StartX = pos.X;
            StartY = pos.Y;
        }

        public override void OnUpdate(Point pos, Size size)
        {
            _lastParentWidth = size.Width;
            _lastParentHeight = size.Height;
            // ドラッグ中は「終点（テキスト側）」だけを更新する
            // 始点 (RelStartX) は OnStart の位置に固定される
            RelX = Normalize(pos.X, size.Width);
            RelY = Normalize(pos.Y, size.Height);

            X = pos.X;
            Y = pos.Y;
        }


        public override bool OnComplete(ICollection<AnnotationBase> parentCollection)
        {
            // ノートはサイズに関わらず残し、編集モードを有効にする
            if (!IsCommitted)
            {
                IsEditing = true;
            }
            return true;
        }

        public override bool OnComplete(Point pos, Size size, string tag)
        {
            if (tag == "Start")
            {
                StartX = pos.X;
                StartY = pos.Y;
            }
            else if (tag == "End")
            {
                X = pos.X;
                Y = pos.Y;
            }
            return true;
        }
    }
}
