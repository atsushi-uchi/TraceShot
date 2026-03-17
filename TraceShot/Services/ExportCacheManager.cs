using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using TraceShot.Models;
using TraceShot.ViewModels;

namespace TraceShot.Services
{
    public class ExportCacheManager
    {
        private readonly Dictionary<Guid, BitmapSource> _imageStore = [];

        public void RegisterCache(Guid id, BitmapSource bitmap)
        {
            if (bitmap != null)
            {
                // 既に freeze されていない場合は、スレッド間共有のために freeze する
                if (!bitmap.IsFrozen) bitmap.Freeze();
                _imageStore[id] = bitmap;
            }
        }

        // キャッシュがあれば返す、なければ null
        public BitmapSource? GetCachedImage(Guid id)
        {
            return _imageStore.TryGetValue(id, out var bitmap) ? bitmap : null;
        }

        // クロップ範囲が変わったら倉庫を空にする（RecServiceから呼ばれる想定）
        public void ClearAll() => _imageStore.Clear();
    }
}
