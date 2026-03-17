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
            return _imageStore.TryGetValue(id, out var bitmap) ? bitmap : null;
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
