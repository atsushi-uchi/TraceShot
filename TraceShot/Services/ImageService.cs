using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using TraceShot.Models;
using Windows.Media.Ocr;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using FontFamily = System.Windows.Media.FontFamily;
using Pen = System.Windows.Media.Pen;
using Point = System.Windows.Point;

// --- クラス内のメンバ変数として、一度生成したビットマップをキャッシュしておくと効率的です ---

namespace TraceShot.Services
{
    public class ImageService
    {
        private static RenderTargetBitmap? _standbyImageCache;

        public static void DebugImageSave(BitmapSource bitmapSource)
        {
            string debugPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ocr_debug.png");
            using (var fileStream = new FileStream(debugPath, FileMode.Create))
            {
                PngBitmapEncoder debugEncoder = new PngBitmapEncoder();
                debugEncoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                debugEncoder.Save(fileStream);
            }
            //Debug.WriteLine($"[OCR Debug] Image saved to: {debugPath}");
        }

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

        public static BitmapSource GeneratePureVideoBitmap(Bookmark bm, VideoSnapshotInfo info)
        {
            // ビデオ本来の解像度
            int w = info.NaturalWidth;
            int h = info.NaturalHeight;

            DrawingVisual dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                // 💡 ここがポイント：VisualBrush を使ってビデオ内容のみを描画
                // これにより、上下左右の黒い余白は完全に除去される
                dc.DrawRectangle(info.VideoBrush, null, new Rect(0, 0, w, h));
            }

            RenderTargetBitmap bmp = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            return bmp;
        }

        public static async Task<string> RecognizeTextFromBitmapSource(BitmapSource bitmapSource)
        {
            try
            {
                // 1. BitmapSource を Png 形式でメモリストリームに保存
                using var ms = new MemoryStream();
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmapSource));
                encoder.Save(ms);
                ms.Seek(0, SeekOrigin.Begin);

                // 2. Windows.Graphics.Imaging のデコーダーで SoftwareBitmap を作成
                // ※ .AsRandomAccessStream() を使うには using System.IO; が必要
                var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(ms.AsRandomAccessStream());
                using var softwareBitmap = await decoder.GetSoftwareBitmapAsync();

                // 3. OCRエンジンの初期化（日本語）
                var engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("ja-JP"));
                if (engine == null) return "OCRエンジンの初期化に失敗しました。";

                // 4. 解析実行
                var ocrResult = await engine.RecognizeAsync(softwareBitmap);

                // 5. 結果を一つの文字列に結合して返す
                return ocrResult.Text;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OCR Recognition Error: {ex.Message}");
                return "";
            }
        }

        public static string BitmapSourceToBase64(BitmapSource bitmapSource)
        {
            if (bitmapSource == null) return string.Empty;

            // 1. エンコーダーを準備（PNGまたはJPEG）
            var encoder = new PngBitmapEncoder(); // 画質重視ならPNG、容量重視ならJpegBitmapEncoder
            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));

            using (var ms = new MemoryStream())
            {
                // 2. メモリストリームに保存
                encoder.Save(ms);
                byte[] imageBytes = ms.ToArray();

                // 3. バイト配列をBase64文字列に変換
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
}
