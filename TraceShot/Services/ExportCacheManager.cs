using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using TraceShot.Models;
using TraceShot.ViewModels;

namespace TraceShot.Services
{
    public class ExportCacheManager
    {
        // GUIDをキーにして、生成済みのViewModel（画像入り）を保存
        private readonly Dictionary<Guid, ExportItemViewModel> _cache = new();

        // 前回のクロップ範囲を保持
        private Rect _lastCropRect;

        public List<ExportItemViewModel> GetOrUpdateCache(VideoSnapshotInfo info)
        {
            var cropRect = RecService.Instance.Evidence.CommonCropRect;
            var bookmarks = RecService.Instance.Bookmarks
                .OrderBy(b => b.Time)
                .ToList();

            // クロップ範囲が変わっていたら、全キャッシュを無効化
            if (_lastCropRect != cropRect)
            {
                _cache.Clear();
                _lastCropRect = cropRect;
            }

            var result = new List<ExportItemViewModel>();

            foreach (var bookmark in bookmarks)
            {
                // キャッシュがあり、かつブックマークが更新されていない（Dirtyでない）なら再利用
                if (_cache.TryGetValue(bookmark.Id, out var cachedItem) && !bookmark.IsDirty)
                {
                    result.Add(cachedItem);
                }
                else
                {
                    // 新しく生成する（ここで実際のキャプチャロジックを呼ぶ）
                    var newItem = CreateExportItem(bookmark, cropRect, info);

                    _cache[bookmark.Id] = newItem;
                    result.Add(newItem);

                    // 出力したのでDirtyフラグを下ろす
                    bookmark.IsDirty = false;
                }
            }

            // キャッシュの中から、既に削除されたブックマークのデータを掃除する（メモリ節約）
            var currentIds = bookmarks.Select(b => b.Id).ToHashSet();
            var removedIds = _cache.Keys.Where(id => !currentIds.Contains(id)).ToList();
            foreach (var id in removedIds) _cache.Remove(id);

            return result;
        }

        private ExportItemViewModel CreateExportItem(Bookmark b, Rect rect, VideoSnapshotInfo info)
        {
            // RecService.SaveImage を呼び出して、注釈込みの画像を生成する
            // scale はプレビュー用なので 0.5 程度で良いかと思います
            var result = RecService.Instance.SaveImage(b, info, scale: 0.5, false);

            if (result == null || result.Value.Bitmap == null)
            {
                // 万が一失敗した時のフォールバック（空の画像など）
                return new ExportItemViewModel(b, null);
            }

            // 生成された画像(Bitmap)と、保存されたパス(Path)を保持
            var newItem = new ExportItemViewModel(b, result.Value.Bitmap);
            newItem.ImagePath = result.Value.Path; // 保存先パスも保持しておくと後で便利です

            return newItem;
        }
    }
}
