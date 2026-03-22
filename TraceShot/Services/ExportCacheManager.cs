using System.Diagnostics;
using System.Windows.Media.Imaging;

namespace TraceShot.Services
{
    public class ExportCacheManager
    {
        private readonly Dictionary<Guid, BitmapSource> _imageStore = [];

        private readonly Dictionary<Guid, int> _itemOrders = [];

        private readonly HashSet<Guid> _excludedIds = [];
        public double LastScale { get; private set; }  = 0;

        public void RegisterCache(Guid id, BitmapSource bitmap, double currentScale)
        {
            if (Math.Abs(LastScale - currentScale) > 0.001)
            {
                _imageStore.Clear();
                LastScale = currentScale;
            }

            if (bitmap != null)
            {
                if (!bitmap.IsFrozen) bitmap.Freeze();
                _imageStore[id] = bitmap;
            }
        }

        public BitmapSource? GetCachedImage(Guid id, double currentScale)
        {
            if (Math.Abs(LastScale - currentScale) > 0.001)
            {
                return null;
            }

            if (_imageStore.TryGetValue(id, out var bitmap))
            {
                return bitmap;
            }

            var entry = RecService.Instance.Entries.FirstOrDefault(e => e.Id == id);
            if (entry != null && !string.IsNullOrEmpty(entry.ImagePath))
            {
                // ファイルの実体が存在するか確認
                if (System.IO.File.Exists(entry.ImagePath))
                {
                    try
                    {
                        // 画像ファイルを読み込んで BitmapSource を生成
                        var uri = new Uri(entry.ImagePath, UriKind.Absolute);
                        var loadedBitmap = new BitmapImage();
                        loadedBitmap.BeginInit();
                        loadedBitmap.CacheOption = BitmapCacheOption.OnLoad; // ファイルロックを避けるためOnLoadを指定
                        loadedBitmap.UriSource = uri;
                        loadedBitmap.EndInit();
                        loadedBitmap.Freeze(); // スレッド間共有を可能にする

                        // 4. メモリキャッシュに登録して返す
                        _imageStore[id] = loadedBitmap;

                        Debug.WriteLine($"画像ファイル読込成功 {entry.ImagePath}");

                        return loadedBitmap;
                    }
                    catch
                    {
                        // 読み込み失敗時は null を返す（再撮影対象にする）
                        return null;
                    }
                }
            }
            return null;
        }

        public void UpdateSelection(Guid id, bool isSelected)
        {
            if (isSelected)
            {
                // チェックを入れたなら、除外リストから消す
                _excludedIds.Remove(id);
            }
            else
            {
                // チェックを外したなら、除外リストに入れる
                _excludedIds.Add(id);
            }
        }

        public void UpdateOrders(IEnumerable<Guid> ids)
        {
            _itemOrders.Clear();
            int index = 0;
            foreach (var id in ids)
            {
                _itemOrders[id] = index++;
            }
        }

        public int GetOrder(Guid id)
        {
            return _itemOrders.TryGetValue(id, out var order) ? order : int.MaxValue;
        }

        public bool IsPreviouslySelected(Guid id)
        {
            return !_excludedIds.Contains(id);
        }

        public bool HasSelectionHistory => _excludedIds.Count > 0;

        public void ClearAll() => _imageStore.Clear();
    }
}
