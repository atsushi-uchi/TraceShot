using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Text.Json.Serialization;
using System.Windows.Input;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    public partial class RectAnnotation : AnnotationBase
    {
        [ObservableProperty] private bool _isDrawing = true;
        [ObservableProperty] private bool _isMasking = false;
        // 外部から代入してもらうためのアクション（非同期対応）
        [JsonIgnore] public Func<RectAnnotation, Task> OcrAction { get; set; }

        public ICommand RunOcrCommand { get; }
        private Point _startPoint;

        public RectAnnotation() : base()
        {
            RunOcrCommand = new RelayCommand(async () =>
            {
                if (OcrAction != null)
                {
                    await OcrAction(this);
                }
            });
        }


        private async Task RequestOcr()
        {
        }

        public override void OnStart(Point pos, Size size)
        {
            base.OnStart(pos, size);

            _startPoint = pos;
        }


        public override void OnUpdate(Point pos, Size size)
        {
            base.OnUpdate(pos, size);

            // 1. 左上座標を特定
            double left = Math.Min(_startPoint.X, pos.X);
            double top = Math.Min(_startPoint.Y, pos.Y);

            // 2. 幅と高さを特定
            double w = Math.Abs(pos.X - _startPoint.X);
            double h = Math.Abs(pos.Y - _startPoint.Y);

            // 3. 絶対座標を更新
            X = left;
            Y = top;
            Width = w;
            Height = h;

            // 4. 相対座標を更新
            RelX = Normalize(left, size.Width);
            RelY = Normalize(top, size.Height);
            RelWidth = Normalize(w, size.Width);
            RelHeight = Normalize(h, size.Height);
        }

        public override bool OnComplete(ICollection<AnnotationBase> parentCollection)
        {
            IsDrawing = false;

            // 小さすぎる場合は自分を削除リストに入れてもらうために false を返す
            return this.Width >= 5 || this.Height >= 5;
        }
    }
}
