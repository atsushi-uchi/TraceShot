using System.IO;
using System.Windows.Media.Imaging;
using Application = System.Windows.Application;

namespace TraceShot.Models
{
    public class ImageData
    {
        public byte[] Bytes { get; set; }
        public double PixelWidth { get; set; }
        public double PixelHeight { get; set; }

        public static ImageData? GetImageData(BitmapSource bitmap)
        {
            if (bitmap == null) return null;
            return Application.Current.Dispatcher.Invoke(() =>
            {
                // 物理ピクセルをそのまま維持し、メタデータを最小限にする
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using var ms = new MemoryStream();
                encoder.Save(ms);

                return new ImageData
                {
                    Bytes = ms.ToArray(),
                    PixelWidth = bitmap.PixelWidth,
                    PixelHeight = bitmap.PixelHeight
                };
            });
        }
    }
}