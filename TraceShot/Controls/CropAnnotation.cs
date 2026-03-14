using CommunityToolkit.Mvvm.ComponentModel;

public enum CropState
{
    Editing,   // 再編集モード（初期値）
    Confirmed  // 確定モード
}

namespace TraceShot.Controls
{
    public partial class CropAnnotation : RectAnnotation
    {
        public CropAnnotation()
        {
            RelWidth = 0.5;
            RelHeight = 0.5;
            RelX = 0.25;
            RelY = 0.25;
        }
    }
}
