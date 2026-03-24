using System.Collections.ObjectModel;
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

        public void RefreshCropOverlay()
        {
            // 1. 既存のオーバーレイ（ガード、クロップ、個別フォーカス用）をクリア
            var overlays = Annotations.Where(a => a is GuardAnnotation || a is CropAnnotation).ToList();
            foreach (var a in overlays) Annotations.Remove(a);

            var localFocus = Annotations.Where(a => a is RectAnnotation rect && rect.IsFocused).FirstOrDefault();
            if (localFocus != null)
            {
                // 個別フォーカス時はガード(GuardAnnotation)を入れない仕様
                var crop = new CropAnnotation
                {
                    // Bookmarkの相対座標をセット
                    RelX = localFocus.RelX,
                    RelY = localFocus.RelY,
                    RelWidth = localFocus.RelWidth,
                    RelHeight = localFocus.RelHeight,
                };

                localFocus.PropertyChanged += (s, e) =>
                {
                    if (s is RectAnnotation ra)
                    {
                        // プロパティ名に応じて同期（RelX, RelY, RelWidth, RelHeight）
                        switch (e.PropertyName)
                        {
                            case nameof(RectAnnotation.RelX): crop.RelX = ra.RelX; break;
                            case nameof(RectAnnotation.RelY): crop.RelY = ra.RelY; break;
                            case nameof(RectAnnotation.RelWidth): crop.RelWidth = ra.RelWidth; break;
                            case nameof(RectAnnotation.RelHeight): crop.RelHeight = ra.RelHeight; break;
                        }
                    }
                };

                Annotations.Add(crop);
                return; // 個別表示時はグローバルを表示しない
            }

            // 3. グローバルフォーカスの表示（従来通り）
            if (RecService.Instance.Evidence.IsCropEnabled)
            {
                if (RecService.Instance.Evidence.CropState == CropState.Editing)
                {
                    Annotations.Add(new GuardAnnotation());
                }

                var crop = new CropAnnotation
                {
                    RelX = RecService.Instance.Evidence.CommonCropRect.X,
                    RelY = RecService.Instance.Evidence.CommonCropRect.Y,
                    RelWidth = RecService.Instance.Evidence.CommonCropRect.Width,
                    RelHeight = RecService.Instance.Evidence.CommonCropRect.Height,
                };

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
        public AnnotationBase StartDrawing<T>(Bookmark bookmark, Point pos, Size size) where T : AnnotationBase, new()
        {
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

            if (SelectedAnnotation is RectAnnotation rect)
            {
                rect.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(RectAnnotation.IsFocused))
                    {
                        RefreshCropOverlay();
                    }
                };
            }

            return SelectedAnnotation;
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
        public void CompleteDrawing(Bookmark? bookmark)
        {
            if (SelectedAnnotation == null) return;

            // 図形自身に「完了していいか？」を判断させる
            bool shouldKeep = SelectedAnnotation.OnComplete(Annotations);

            if (!shouldKeep)
            {
                Annotations.Remove(SelectedAnnotation);
                bookmark?.Annotations.Remove(SelectedAnnotation);
            }

            SelectedAnnotation = null;
        }

        public void CompleteDrawing(AnnotationBase annotation, Point pos, Size size, string tag)
        {
            annotation.OnComplete(pos, size, tag);
            SelectedAnnotation = null;
        }

        public void ClearAll() => Annotations.Clear();
    }
}