using CommunityToolkit.Mvvm.ComponentModel;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.ObjectModel;
using System.Security.Policy;
using System.Windows;
using TraceShot.Models;
using TraceShot.Services;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    public partial class AnnotationManager
    {
        // 確定した注釈のリスト（これをXAMLのItemsControlにバインドする）
        public ObservableCollection<AnnotationBase> Annotations { get; } = new();

        public AnnotationBase? SelectedAnnotation { get; private set; }

        private Point _startPoint;

        // クロップ範囲の同期
        // 現在表示中のブックマークに共通枠を適用する
        public void RefreshCropOverlay()
        {
            // 既存の CropAnnotation を一旦クリア
            var existing = Annotations.OfType<CropAnnotation>().ToList();
            foreach (var a in existing) Annotations.Remove(a);

            if (RecService.Instance.Evidence.IsCropEnabled)
            {
                // 共通座標から新しい Annotation を生成
                var crop = new CropAnnotation
                {
                    RelX = RecService.Instance.Evidence.CommonCropRect.X,
                    RelY = RecService.Instance.Evidence.CommonCropRect.Y,
                    RelWidth = RecService.Instance.Evidence.CommonCropRect.Width,
                    RelHeight = RecService.Instance.Evidence.CommonCropRect.Height,
                };
                RecService.Instance.Evidence.CropState = CropState.Confirmed;

                // 枠が動かされた時に、共通データに書き戻すイベントを購読
                crop.PropertyChanged += (s, e) => {
                    if (s is CropAnnotation ca)
                    {
                        RecService.Instance.Evidence.CommonCropRect = new Rect(ca.RelX, ca.RelY, ca.RelWidth, ca.RelHeight);
                    }
                };

                Annotations.Add(crop);
            }
        }
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
            RefreshCropOverlay();
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