namespace TraceShot.Controls
{
    public partial class GuardAnnotation : AnnotationBase
    {
        public GuardAnnotation()
        {
            // 常に画面全体を覆うので、座標やサイズは固定、または親に追従させる
            RelX = 0;
            RelY = 0;
            RelWidth = 1.0;
            RelHeight = 1.0;
            ZIndex = 5000; // 注釈(0)より高く、クロップ(10000)より低い値
        }
    }
}
