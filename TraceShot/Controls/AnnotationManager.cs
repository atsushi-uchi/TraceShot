using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using TraceShot.Models;
using TraceShot.Services;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    public partial class AnnotationManager : ObservableObject
    {
        private readonly Stack<IUndoableAction> _undoStack = new();
        private readonly Stack<IUndoableAction> _redoStack = new();

        public bool CanUndo => _undoStack.Count > 0;
        public bool CanRedo => _redoStack.Count > 0;

        /// <summary>
        /// 状態変更前のスナップショットを撮り、変更を実行してスタックに積みます。
        /// </summary>
        public void ExecuteRectStateChange(IEnumerable<RectAnnotation> allRects, RectAnnotation target, string mode)
        {
            // 1. 変更前の全状態を記録
            var before = allRects.Select(r => (r, r.IsFocused, r.IsMasking)).ToList();

            // 2. 実際の変更処理（ここで RectAnnotation 内の partial メソッドによる排他ロジックが走る）
            if (mode == "Focus")
            {
                target.IsFocused = !target.IsFocused;
            }
            else if (mode == "Masking")
            {
                target.IsMasking = !target.IsMasking;
            }

            // 3. 変更後の全状態を記録
            var after = allRects.Select(r => (r, r.IsFocused, r.IsMasking)).ToList();

            // 4. 変化があった場合のみスタックに積む
            bool changed = false;
            for (int i = 0; i < before.Count; i++)
            {
                if (before[i].IsFocused != after[i].IsFocused || before[i].IsMasking != after[i].IsMasking)
                {
                    changed = true;
                    break;
                }
            }

            if (changed)
            {
                PushAction(new RectStateAction(this, before, after));
            }
        }

        public void PushBoolPropertyAction(AnnotationBase annotation, string propertyName, bool before, bool after, Action? onChanged = null)
        {
            if (annotation == null || before == after) return;

            PushAction(new UpdateBoolPropertyAction(annotation, propertyName, before, after, onChanged));
        }

        public void PushTextUpdateAction(NoteAnnotation note)
        {
            if (note == null || note.OriginText == note.Text) return;

            PushAction(new UpdateTextAction(note, note.OriginText, note.Text));

            note.OriginText = note.Text;
        }

        private void PushAction(IUndoableAction action)
        {
            _undoStack.Push(action);
            _redoStack.Clear();
            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }

        public void Undo()
        {
            if (!CanUndo) return;

            var currentSelected = SelectedAnnotation;
            SelectedAnnotation = null; // 一時的に選択解除してUIの干渉を遮断

            var act = _undoStack.Pop();
            act.Undo();
            _redoStack.Push(act);

            RefreshCropOverlay();
            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }

        public void Redo()
        {
            if (!CanRedo) return;

            var currentSelected = SelectedAnnotation;
            SelectedAnnotation = null; // 一時的に選択解除してUIの干渉を遮断

            var act = _redoStack.Pop();
            act.Redo();
            _undoStack.Push(act);

            RefreshCropOverlay();
            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }

        /// <summary>
        /// 注釈の更新前後の状態を保持するクラス
        /// </summary>
        public class UpdateState
        {
            public double RelX;
            public double RelY;
            public double RelWidth;
            public double RelHeight;
            public double? RelStartX;
            public double? RelStartY;
        }

        // 操作開始時の状態を一時保存する辞書
        private readonly Dictionary<Guid, UpdateState> _updateStartState = new();

        /// <summary>
        /// ドラッグ開始時の状態を記録します。UI側のDragStartedなどで呼び出してください。
        /// </summary>
        public void RecordUpdateStart(AnnotationBase annotation)
        {
            if (annotation == null) return;

            var s = new UpdateState
            {
                RelX = annotation.RelX,
                RelY = annotation.RelY,
                RelWidth = annotation.RelWidth,
                RelHeight = annotation.RelHeight
            };

            if (annotation is NoteAnnotation note)
            {
                s.RelStartX = note.RelStartX;
                s.RelStartY = note.RelStartY;
            }

            _updateStartState[annotation.Id] = s;
        }

        /// <summary>
        /// 記録された開始状態を取得し、辞書から削除します。
        /// </summary>
        public bool TryConsumeUpdateStart(Guid id, out UpdateState? start)
        {
            if (_updateStartState.TryGetValue(id, out var s))
            {
                _updateStartState.Remove(id);
                start = s;
                return true;
            }
            start = null;
            return false;
        }

        /// <summary>
        /// 更新アクションをスタックに積みます。DragCompletedなどで呼び出してください。
        /// </summary>
        public void PushUpdateAction(AnnotationBase annotation, UpdateState before, UpdateState after)
        {
            if (annotation == null || before == null || after == null) return;

            // 開始時と終了時が全く同じ（1ピクセルも動いていない）なら、スタックに積まない
            if (before.RelX == after.RelX &&
                    before.RelY == after.RelY &&
                    before.RelWidth == after.RelWidth &&
                    before.RelHeight == after.RelHeight &&
                    before.RelStartX == after.RelStartX &&
                    before.RelStartY == after.RelStartY)
            {
                Debug.WriteLine("変化がないためスタックをスキップしました");
                return;
            }
            Debug.WriteLine($"PushUpdateAction {annotation.Id}");

            // 移動距離が極端に小さい場合は無視するロジックを入れても良い
            PushAction(new UpdateAnnotationAction(this, annotation, before, after));
        }

        private bool IsSameState(UpdateState a, UpdateState b)
        {
            bool basicSame = a.RelX == b.RelX &&
                             a.RelY == b.RelY &&
                             a.RelWidth == b.RelWidth &&
                             a.RelHeight == b.RelHeight;

            bool noteSame = a.RelStartX == b.RelStartX &&
                            a.RelStartY == b.RelStartY;

            return basicSame && noteSame;
        }

        /// <summary>
        /// 座標更新用のUndo/Redoアクション
        /// </summary>
        internal class UpdateAnnotationAction : IUndoableAction
        {
            private readonly AnnotationManager _mgr;
            private readonly AnnotationBase _annotation;
            private readonly UpdateState _before;
            private readonly UpdateState _after;

            public UpdateAnnotationAction(AnnotationManager mgr, AnnotationBase annotation, UpdateState before, UpdateState after)
            {
                _mgr = mgr;
                _annotation = annotation;
                _before = before;
                _after = after;
            }

            public void Undo() => ApplyState(_before);
            public void Redo() => ApplyState(_after);

            private void ApplyState(UpdateState s)
            {
                _annotation.RelX = s.RelX;
                _annotation.RelY = s.RelY;
                _annotation.RelWidth = s.RelWidth;
                _annotation.RelHeight = s.RelHeight;

                if (_annotation is NoteAnnotation note)
                {
                    if (s.RelStartX.HasValue) note.RelStartX = s.RelStartX.Value;
                    if (s.RelStartY.HasValue) note.RelStartY = s.RelStartY.Value;
                }

                _mgr.RefreshCropOverlay();
            }
        }

        public ObservableCollection<AnnotationBase> Annotations { get; } = new();
        public AnnotationBase? CreatingAnnotation { get; private set; }

        private AnnotationBase? _selectedAnnotation;
        public AnnotationBase? SelectedAnnotation
        {
            get => _selectedAnnotation;
            set
            {
                if (_selectedAnnotation == value) return;
                if (_selectedAnnotation != null)
                {
                    _selectedAnnotation.IsSelected = false;
                }
                if (SetProperty(ref _selectedAnnotation, value))
                {
                    if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = true;
                }
            }
        }

        public void RefreshCropOverlay()
        {
            var overlays = Annotations.Where(a => a is GuardAnnotation || a is CropAnnotation).ToList();
            foreach (var a in overlays) Annotations.Remove(a);

            var localFocus = Annotations.Where(a => a is RectAnnotation rect && rect.IsFocused).FirstOrDefault();
            if (localFocus != null)
            {
                var crop = new CropAnnotation
                {
                    RelX = localFocus.RelX,
                    RelY = localFocus.RelY,
                    RelWidth = localFocus.RelWidth,
                    RelHeight = localFocus.RelHeight,
                };

                localFocus.PropertyChanged += (s, e) =>
                {
                    if (s is RectAnnotation ra)
                    {
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
                return;
            }

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
            Annotations.Clear();
            if (bookmark != null && bookmark.Annotations != null)
            {
                foreach (var item in bookmark.Annotations) Annotations.Add(item);
            }
            RefreshCropOverlay();
        }

        public void Remove(Bookmark? bookmark, AnnotationBase target)
        {
            if (target is null) return;
            if (bookmark != null && bookmark.Annotations.Contains(target))
            {
                bookmark.Annotations.Remove(target);
                PushAction(new RemoveAnnotationAction(this, bookmark, target));
            }
            if (Annotations.Contains(target))
            {
                Annotations.Remove(target);
                if (CreatingAnnotation == target) CreatingAnnotation = null;
            }
        }

        public void Select(AnnotationBase? target)
        {
            Annotations.ToList().ForEach(a => a.IsSelected = false);
            CreatingAnnotation = target;
            if (CreatingAnnotation != null) CreatingAnnotation.IsSelected = true;
        }

        public AnnotationBase StartDrawing<T>(Bookmark bookmark, Point pos, Size size) where T : AnnotationBase, new()
        {
            CreatingAnnotation = new T { X = pos.X, Y = pos.Y, Width = 0, Height = 0 };
            CreatingAnnotation.OnStart(pos, size);
            bookmark.Annotations.Add(CreatingAnnotation);
            Annotations.Add(CreatingAnnotation);
            PushAction(new AddAnnotationAction(this, bookmark, CreatingAnnotation));
            if (CreatingAnnotation is RectAnnotation rect)
            {
                rect.PropertyChanged += (s, e) => { if (e.PropertyName == nameof(RectAnnotation.IsFocused)) RefreshCropOverlay(); };
            }
            return CreatingAnnotation;
        }

        public void UpdateDrawing(Point pos, Size size)
        {
            if (CreatingAnnotation == null) return;
            CreatingAnnotation.OnUpdate(pos, size);
        }

        public void CompleteDrawing(Bookmark? bookmark)
        {
            if (CreatingAnnotation == null) return;

            bool shouldKeep = CreatingAnnotation.OnComplete(Annotations);

            if (CreatingAnnotation.RelWidth < 0.001 && CreatingAnnotation.RelHeight < 0.001)
            {
                shouldKeep = false;
            }

            if (!shouldKeep)
            {
                Annotations.Remove(CreatingAnnotation);
                bookmark?.Annotations.Remove(CreatingAnnotation);
                if (CanUndo && _undoStack.Peek() is AddAnnotationAction addAct && addAct.Annotation == CreatingAnnotation)
                {
                    _undoStack.Pop();
                    OnPropertyChanged(nameof(CanUndo));
                }
            }
            SelectedAnnotation = CreatingAnnotation;
            CreatingAnnotation = null;
        }

        public void CompleteDrawing(AnnotationBase annotation, Point pos, Size size, string tag)
        {
            annotation.OnComplete(pos, size, tag);
            SelectedAnnotation = annotation;
            CreatingAnnotation = null;
        }

        public void ClearAll() => Annotations.Clear();

        public void AddPastedAnnotation(Bookmark bookmark, AnnotationBase annotation)
        {
            if (annotation is RectAnnotation rect && rect.IsFocused)
            {
                var list = Annotations.OfType<RectAnnotation>().ToList();
                foreach (var other in list) if (other != rect) other.IsFocused = false;
            }
            bookmark.Annotations.Add(annotation);
            Annotations.Add(annotation);
            SelectedAnnotation = annotation;
            PushAction(new AddAnnotationAction(this, bookmark, annotation));
            if (annotation is RectAnnotation r)
            {
                r.PropertyChanged += (s, e) => { if (e.PropertyName == nameof(RectAnnotation.IsFocused)) RefreshCropOverlay(); };
            }
            RefreshCropOverlay();
        }
    }

    internal interface IUndoableAction
    {
        void Undo();
        void Redo();
    }

    internal class AddAnnotationAction : IUndoableAction
    {
        private readonly AnnotationManager _mgr;
        private readonly Bookmark _bookmark;
        public readonly AnnotationBase Annotation;

        public AddAnnotationAction(AnnotationManager mgr, Bookmark bookmark, AnnotationBase annotation)
        {
            _mgr = mgr; _bookmark = bookmark; Annotation = annotation;
        }

        public void Undo()
        {
            if (_bookmark.Annotations.Contains(Annotation)) _bookmark.Annotations.Remove(Annotation);
            if (_mgr.Annotations.Contains(Annotation)) _mgr.Annotations.Remove(Annotation);
        }

        public void Redo()
        {
            if (!_bookmark.Annotations.Contains(Annotation)) _bookmark.Annotations.Add(Annotation);
            if (!_mgr.Annotations.Contains(Annotation)) _mgr.Annotations.Add(Annotation);
        }
    }

    internal class RemoveAnnotationAction : IUndoableAction
    {
        private readonly AnnotationManager _mgr; private readonly Bookmark _bookmark; private readonly AnnotationBase _annotation;
        public RemoveAnnotationAction(AnnotationManager mgr, Bookmark bookmark, AnnotationBase annotation)
        {
            _mgr = mgr; _bookmark = bookmark; _annotation = annotation;
        }
        public void Undo()
        {
            if (!_bookmark.Annotations.Contains(_annotation)) _bookmark.Annotations.Add(_annotation);
            if (!_mgr.Annotations.Contains(_annotation)) _mgr.Annotations.Add(_annotation);
        }
        public void Redo()
        {
            if (_bookmark.Annotations.Contains(_annotation)) _bookmark.Annotations.Remove(_annotation);
            if (_mgr.Annotations.Contains(_annotation)) _mgr.Annotations.Remove(_annotation);
        }
    }

    internal class UpdateTextAction : IUndoableAction
    {
        private readonly NoteAnnotation _note;
        private readonly string _oldText;
        private readonly string _newText;

        public UpdateTextAction(NoteAnnotation note, string oldText, string newText)
        {
            _note = note;
            _oldText = oldText;
            _newText = newText;
        }

        public void Undo() => _note.Text = _oldText;
        public void Redo() => _note.Text = _newText;
    }

    internal class UpdateBoolPropertyAction : IUndoableAction
    {
        private readonly AnnotationBase _annotation;
        private readonly string _propertyName;
        private readonly bool _before;
        private readonly bool _after;
        private readonly Action? _onChanged;

        public UpdateBoolPropertyAction(AnnotationBase annotation, string propertyName, bool before, bool after, Action? onChanged = null)
        {
            _annotation = annotation;
            _propertyName = propertyName;
            _before = before;
            _after = after;
            _onChanged = onChanged;
        }

        public void Undo() => Apply(_before);
        public void Redo() => Apply(_after);

        private void Apply(bool value)
        {
            var prop = _annotation.GetType().GetProperty(_propertyName);
            prop?.SetValue(_annotation, value);
            _onChanged?.Invoke(); // 必要に応じて描画更新などを走らせる
        }
    }

    internal class RectStateAction : IUndoableAction
    {
        private readonly AnnotationManager _mgr;
        // (対象矩形, Focus状態, Masking状態) のリスト
        private readonly List<(RectAnnotation Rect, bool Focused, bool Masking)> _before;
        private readonly List<(RectAnnotation Rect, bool Focused, bool Masking)> _after;

        public RectStateAction(AnnotationManager mgr,
                               List<(RectAnnotation, bool, bool)> before,
                               List<(RectAnnotation, bool, bool)> after)
        {
            _mgr = mgr;
            _before = before;
            _after = after;
        }

        public void Undo() => Apply(_before);
        public void Redo() => Apply(_after);

        private void Apply(List<(RectAnnotation Rect, bool Focused, bool Masking)> states)
        {
            foreach (var item in states)
            {
                // Note: Setter内のロジック(OnIsFocusedChangedなど)が走りますが、
                // 全体の整合性が取れた状態(Snapshot)を順次適用するため、最終的に正しい状態に戻ります。
                item.Rect.IsFocused = item.Focused;
                item.Rect.IsMasking = item.Masking;
            }
            _mgr.RefreshCropOverlay();
        }
    }
}