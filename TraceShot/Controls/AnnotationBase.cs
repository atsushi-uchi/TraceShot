using CommunityToolkit.Mvvm.ComponentModel;
using System.Text.Json.Serialization;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    [JsonDerivedType(typeof(RectAnnotation), typeDiscriminator: "rect")]
    [JsonDerivedType(typeof(NoteAnnotation), typeDiscriminator: "note")]
    public abstract partial class AnnotationBase : ObservableObject
    {
        // 0.0 ～ 1.0 の相対値
        [ObservableProperty] private double _relX;
        [ObservableProperty] private double _relY;
        [ObservableProperty] private double _relWidth;
        [ObservableProperty] private double _relHeight;

        [ObservableProperty] private double _x;
        [ObservableProperty] private double _y;
        [ObservableProperty] private double _width;
        [ObservableProperty] private double _height;

        [ObservableProperty] private bool _isVisible = true;
        [ObservableProperty] private bool _isSelected;

        protected Point _startPoint;

        public Guid Id { get; } = Guid.NewGuid();

        public virtual void Delete() { }

        public virtual void OnStart(Point pos, Size size)
        {
            _startPoint = pos;

            // 基本的な位置を相対座標で保持
            X = pos.X;
            Y = pos.Y;
            RelX = Normalize(pos.X, size.Width);
            RelY = Normalize(pos.Y, size.Height);

            Width = 0;
            Height = 0;
            RelWidth = 0;
            RelHeight = 0;
        }

        public virtual void OnUpdate(Point pos, Size size)
        {
        }

        public virtual bool OnComplete(ICollection<AnnotationBase> parentCollection)
        {
            return true;
        }
        public double Normalize(double pos, double total) => total > 0 ? pos / total : 0;
    }
}