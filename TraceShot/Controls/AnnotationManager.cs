using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.ObjectModel;
using TraceShot.Models;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    public class AnnotationManager
    {
        // 確定した注釈のリスト（これをXAMLのItemsControlにバインドする）
        public ObservableCollection<AnnotationBase> Annotations { get; } = new();

        public AnnotationBase? SelectedAnnotation { get; private set; }

        private Point _startPoint;

        public void LoadAnnotationsFromBookmark(Bookmark bookmark)
        {
            // 現在画面に表示されている注釈をすべてクリア
            Annotations.Clear();

            if (bookmark != null && bookmark.Annotations != null)
            {
                // ブックマークが持っている注釈を、表示用リストにコピー
                foreach (var item in bookmark.Annotations)
                {
                    Annotations.Add(item);
                }
            }
        }

        public void Remove(Bookmark? bookmark, AnnotationBase target)
        {
            if (target is null) return;

            if (bookmark != null && bookmark.Annotations.Contains(target))
            {
                bookmark.Annotations.Remove(target);
            }
            if (Annotations.Contains(target))
            {
                Annotations.Remove(target);
                // 削除したものが選択中だった場合は解除
                if (SelectedAnnotation == target)
                {
                    SelectedAnnotation = null;
                }
            }
        }

        public void Select(AnnotationBase? target)
        {
            // 1. 今まで選んでいたものの選択を解除
            Annotations.ToList().ForEach(a => a.IsSelected = false);

            // 2. 新しいものを選択
            SelectedAnnotation = target;

            if (SelectedAnnotation != null)
            {
                SelectedAnnotation.IsSelected = true;
            }
        }

        /// <summary>
        /// 注釈の描画を開始する（MouseDownで呼ぶ）
        /// </summary>
        public void StartDrawing<T>(Bookmark bookmark, Point pos, Size size) where T : AnnotationBase, new()
        {
            _startPoint = pos;
            SelectedAnnotation = new T
            {
                X = pos.X,
                Y = pos.Y,
                Width = 0,
                Height = 0
            };

            SelectedAnnotation.OnStart(pos, size);

            bookmark.Annotations.Add(SelectedAnnotation);

            Annotations.Add(SelectedAnnotation);
        }

        /// <summary>
        /// ドラッグに合わせてサイズを更新する（MouseMoveで呼ぶ）
        /// </summary>
        public void UpdateDrawing(Point pos, Size size)
        {
            if (SelectedAnnotation == null) return;

            SelectedAnnotation.OnUpdate(pos, size);
        }

        /// <summary>
        /// 描画を確定させる（MouseUpで呼ぶ）
        /// </summary>
        public void CompleteDrawing()
        {
            if (SelectedAnnotation == null) return;

            // 図形自身に「完了していいか？」を判断させる
            bool shouldKeep = SelectedAnnotation.OnComplete(Annotations);

            if (!shouldKeep)
            {
                Annotations.Remove(SelectedAnnotation);
            }

            SelectedAnnotation = null;
        }

        /// <summary>
        /// 全ての注釈をクリアする
        /// </summary>
        public void ClearAll() => Annotations.Clear();
    }
}