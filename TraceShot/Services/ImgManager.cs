using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Globalization;
using System.Windows;
using Color = System.Windows.Media.Color;
using Point = System.Windows.Point;
using Pen = System.Windows.Media.Pen;
using FontFamily = System.Windows.Media.FontFamily;
using Brushes = System.Windows.Media.Brushes;

// --- クラス内のメンバ変数として、一度生成したビットマップをキャッシュしておくと効率的です ---

namespace TraceShot.Services
{
    public class ImgManager
    {
        private static RenderTargetBitmap? _standbyImageCache;

        /// <summary>
        /// 録画準備完了（スタンドバイ）のイメージを生成し、PreviewImageに表示します。
        /// </summary>
        public static RenderTargetBitmap? GetReadyStandardImage(double pixelsPerDip = 1.0)
        {
            // キャッシュがあればそれを利用（毎回生成しない）
            if (_standbyImageCache != null)
            {
                return _standbyImageCache;
            }

            // --- イメージの生成処理 ---

            // 1. 生成するイメージのサイズ（プレビューの解像度に合わせて調整してください）
            int width = 1920;
            int height = 1080;
            double dpi = 96;

            // 2. 描画用の Visual を作成
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                // 💡 a) 背景（少しグラデーションをかけるとプロっぽくなります）
                var backgroundBrush = new LinearGradientBrush(
                    Color.FromRgb(30, 30, 30), // 濃いグレー
                    Color.FromRgb(10, 10, 10), // ほぼ黒
                    new Point(0, 0), new Point(0, 1) // 縦方向のグラデーション
                );
                dc.DrawRectangle(backgroundBrush, null, new Rect(0, 0, width, height));

                // 💡 b) 「STANDBY」の文字（中央に配置）
                var textTypeface = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Bold, FontStretches.Normal);
                var formattedText = new FormattedText(
                    "STANDBY",
                    CultureInfo.InvariantCulture,
                    System.Windows.FlowDirection.LeftToRight,
                    textTypeface,
                    120, // フォントサイズ
                    Brushes.Gray, // 文字の色
                    pixelsPerDip // DPI対応
                );
                formattedText.TextAlignment = TextAlignment.Center; // 中央揃え

                // 文字を描画（画面の中央）
                dc.DrawText(formattedText, new Point(width / 2, (height / 2) - (formattedText.Height / 2)));

                // 💡 c) 装飾（の下に薄いラインを入れるなど）
                var linePen = new Pen(new SolidColorBrush(Color.FromArgb(100, 255, 255, 255)), 2); // 半透明の白
                dc.DrawLine(linePen, new Point(width * 0.3, height * 0.6), new Point(width * 0.7, height * 0.6));
            }

            // 3. DrawingVisual をビットマップにレンダリング
            _standbyImageCache = new RenderTargetBitmap(width, height, dpi, dpi, PixelFormats.Pbgra32);
            _standbyImageCache.Render(visual);

            // 💡 キャッシュのフリーズ（スレッド間共有を可能にし、パフォーマンスを向上させる）
            _standbyImageCache.Freeze();

            // 4. PreviewImage にセット
            return _standbyImageCache;
        }
    }
}
