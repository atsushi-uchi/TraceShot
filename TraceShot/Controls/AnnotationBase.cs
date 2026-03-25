using CommunityToolkit.Mvvm.ComponentModel;
using System.Text.Json.Serialization;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    [JsonDerivedType(typeof(GuardAnnotation), typeDiscriminator: "guard")]
    [JsonDerivedType(typeof(CropAnnotation), typeDiscriminator: "crop")]
    [JsonDerivedType(typeof(RectAnnotation), typeDiscriminator: "rect")]
    [JsonDerivedType(typeof(NoteAnnotation), typeDiscriminator: "note")]
    public abstract partial class AnnotationBase : ObservableObject
    {
        [ObservableProperty] private int _zIndex = 0;
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
        [JsonIgnore]
        [ObservableProperty] private bool _isSelected;

        public Guid Id { get; } = Guid.NewGuid();

        public virtual void Delete() { }

        public virtual void OnStart(Point pos, Size size)
        {
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

        public virtual bool OnComplete(Point pos, Size size, string tag)
        {
            return true;
        }

        public void UpdateRelativeMetrics(Size canvasSize)
        {
            if (canvasSize.Width > 0 && canvasSize.Height > 0)
            {
                RelX = X / canvasSize.Width;
                RelY = Y / canvasSize.Height;
                RelWidth = Width / canvasSize.Width;
                RelHeight = Height / canvasSize.Height;
            }
        }

        public void UpdateAbsoluteMetrics(Size canvasSize)
        {
            if (canvasSize.Width > 0 && canvasSize.Height > 0)
            {
                X = RelX * canvasSize.Width;
                Y = RelY * canvasSize.Height;
                Width = RelWidth * canvasSize.Width;
                Height = RelHeight * canvasSize.Height;
            }
        }

        public double Normalize(double pos, double total) => total > 0 ? pos / total : 0;
    }
}