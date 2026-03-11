using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TraceShot.Models
{
    public class VideoSnapshotInfo
    {
        public int NaturalWidth { get; }
        public int NaturalHeight { get; }
        public VisualBrush VideoBrush { get; }
        public double DpiScale { get; }
        public double ActualViewWidth { get; }
        public double ActualViewHeight { get; }

        public VideoSnapshotInfo(MediaElement videoPlayer)
        {
            // その時点の解像度を保持
            NaturalWidth = videoPlayer.NaturalVideoWidth;
            NaturalHeight = videoPlayer.NaturalVideoHeight;
            ActualViewWidth = videoPlayer.ActualWidth;
            ActualViewHeight = videoPlayer.ActualHeight;

            // VisualBrushを作成。元のコントロールへの依存を最小限にするため設定を行う
            VideoBrush = new VisualBrush(videoPlayer)
            {
                Stretch = Stretch.Uniform,
                ViewboxUnits = BrushMappingMode.RelativeToBoundingBox,
                Viewbox = new Rect(0, 0, 1, 1)
            };

            // DPI情報の取得（FormattedTextなどで使用）
            DpiScale = VisualTreeHelper.GetDpi(videoPlayer).PixelsPerDip;
        }
    }
}
