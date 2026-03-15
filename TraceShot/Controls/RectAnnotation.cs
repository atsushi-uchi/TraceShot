using CommunityToolkit.Mvvm.ComponentModel;
using System.Diagnostics;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    public partial class RectAnnotation : AnnotationBase
    {
        [ObservableProperty]
        private bool _isDrawing = true;

        protected Point _startPoint;

        public override void OnStart(Point pos, Size size)
        {
            base.OnStart(pos, size);

            _startPoint = pos;
        }


        public override void OnUpdate(Point pos, Size size)
        {
            Debug.WriteLine($"RectAnnotation OnUpdate {pos} {size}");

            //base.OnUpdate(pos, size);

            //// 1. 左上座標を特定
            //double left = Math.Min(_startPoint.X, pos.X);
            //double top = Math.Min(_startPoint.Y, pos.Y);

            //// 2. 幅と高さを特定
            //double w = Math.Abs(pos.X - _startPoint.X);
            //double h = Math.Abs(pos.Y - _startPoint.Y);

            //// 3. 絶対座標を更新
            //X = left;
            //Y = top;
            //Width = w;
            //Height = h;

            //// 4. 相対座標を更新
            //RelX = Normalize(left, size.Width);
            //RelY = Normalize(top, size.Height);
            //RelWidth = Normalize(w, size.Width);
            //RelHeight = Normalize(h, size.Height);
        }

        public override bool OnComplete(ICollection<AnnotationBase> parentCollection)
        {
            Debug.WriteLine($"RectAnnotation OnComplete {parentCollection}");

            IsDrawing = false;

            // 小さすぎる場合は自分を削除リストに入れてもらうために false を返す
            return this.Width >= 5 || this.Height >= 5;
        }
    }
}
