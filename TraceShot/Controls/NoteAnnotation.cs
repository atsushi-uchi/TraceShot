using CommunityToolkit.Mvvm.ComponentModel;
using System.Diagnostics;
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

        public override void OnStart(Point pos, Size size)
        {
            base.OnStart(pos, size);

            RelStartX = Normalize(pos.X, size.Width);
            RelStartY = Normalize(pos.Y, size.Height);

            StartX = pos.X;
            StartY = pos.Y;
        }

        public override void OnUpdate(Point pos, Size size)
        {
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
