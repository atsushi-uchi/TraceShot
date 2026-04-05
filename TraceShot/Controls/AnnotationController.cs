using System.Windows;
using System.Windows.Controls.Primitives;
using TraceShot.Models;
using TraceShot.ViewModels;
using Point = System.Windows.Point;
using Size = System.Windows.Size;

namespace TraceShot.Controls
{
    internal class AnnotationController
    {
        private readonly MainViewModel _data;
        private readonly Func<AppViewMode> _getMode;
        private readonly Func<FrameworkElement> _getVideoElement; // VideoPlayer or RescueImage depending on mode
        private readonly Func<System.Windows.Controls.ListBox> _getBookmarkListBox;

        public AnnotationController(MainViewModel data, Func<AppViewMode> getMode, Func<FrameworkElement> getVideoElement, Func<System.Windows.Controls.ListBox> getBookmarkListBox)
        {
            _data = data;
            _getMode = getMode;
            _getVideoElement = getVideoElement;
            _getBookmarkListBox = getBookmarkListBox;
        }

        public void OnAnnotation_DragStarted(object? sender, DragStartedEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is AnnotationBase annotation)
            {
                _data.AnnotationManager.RecordUpdateStart(annotation);
            }
        }

        public void OnAnnotation_DragCompleted(object? sender, DragCompletedEventArgs e)
        {
            if (sender is not Thumb thumb || thumb.DataContext is not AnnotationBase annotation) return;

            var referenceElement = _getVideoElement();
            if (referenceElement == null) return;

            var transform = thumb.TransformToVisual(referenceElement);
            Point currentPos = transform.Transform(new Point(0, 0));

            var tag = thumb.Tag?.ToString() ?? string.Empty;
            if (annotation is NoteAnnotation note && tag == "End")
            {
                currentPos.X += note.ActualTextWidth / 2.0;
                currentPos.Y += note.ActualTextHeight / 2.0;
            }

            var actualSize = new Size(referenceElement.ActualWidth, referenceElement.ActualHeight);

            _data.AnnotationManager.CompleteDrawing(annotation, currentPos, actualSize, tag);

            if (_data.AnnotationManager.TryConsumeUpdateStart(annotation.Id, out var before))
            {
                var after = new AnnotationManager.UpdateState
                {
                    RelX = annotation.RelX,
                    RelY = annotation.RelY,
                    RelWidth = annotation.RelWidth,
                    RelHeight = annotation.RelHeight,
                };

                if (annotation is NoteAnnotation n)
                {
                    after.RelStartX = n.RelStartX;
                    after.RelStartY = n.RelStartY;
                }

                _data.AnnotationManager.PushUpdateAction(annotation, before!, after, _data.SelectedItem);
            }

            // Dirty flag update
            if (annotation is CropAnnotation)
            {
                foreach (var bm in _data.Recorder.Entries) bm.Modified();
                _data.AnnotationManager.RefreshCropOverlay();
            }
            else if (_getBookmarkListBox().SelectedItem is Bookmark bookmark)
            {
                bookmark.Modified();
            }
        }

        public void OnNoteStart_DragDelta(object? sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is NoteAnnotation note)
            {
                var el = _getVideoElement();
                if (el == null) return;

                note.RelStartX += e.HorizontalChange / el.ActualWidth;
                note.RelStartY += e.VerticalChange / el.ActualHeight;
            }
        }

        public void OnNoteText_DragDelta(object? sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is NoteAnnotation note)
            {
                if (note.IsEditing) return;
                var el = _getVideoElement();
                if (el == null) return;

                note.RelX += e.HorizontalChange / el.ActualWidth;
                note.RelY += e.VerticalChange / el.ActualHeight;

                note.RelX = Math.Clamp(note.RelX, 0, 1);
                note.RelY = Math.Clamp(note.RelY, 0, 1);
            }
        }

        public void OnMoveThumb_DragDelta(object? sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                var el = _getVideoElement();
                if (el == null) return;

                rect.RelX += e.HorizontalChange / el.ActualWidth;
                rect.RelY += e.VerticalChange / el.ActualHeight;
            }
        }

        public void OnResize_DragDelta(object? sender, DragDeltaEventArgs e)
        {
            if (sender is not Thumb t || t.DataContext is not RectAnnotation rect) return;

            var el = _getVideoElement();
            if (el == null) return;

            double deltaX = e.HorizontalChange / el.ActualWidth;
            double deltaY = e.VerticalChange / el.ActualHeight;

            switch (t.Tag?.ToString())
            {
                case "Left":
                    rect.RelX += deltaX;
                    rect.RelWidth -= deltaX;
                    break;
                case "Right":
                    rect.RelWidth += deltaX;
                    break;
                case "Top":
                    rect.RelY += deltaY;
                    rect.RelHeight -= deltaY;
                    break;
                case "Bottom":
                    rect.RelHeight += deltaY;
                    break;
                case "TopLeft":
                    rect.RelX += deltaX;
                    rect.RelWidth -= deltaX;
                    rect.RelY += deltaY;
                    rect.RelHeight -= deltaY;
                    break;
                case "TopRight":
                    rect.RelWidth += deltaX;
                    rect.RelY += deltaY;
                    rect.RelHeight -= deltaY;
                    break;
                case "BottomLeft":
                    rect.RelX += deltaX;
                    rect.RelWidth -= deltaX;
                    rect.RelHeight += deltaY;
                    break;
                case "BottomRight":
                    rect.RelWidth += deltaX;
                    rect.RelHeight += deltaY;
                    break;
            }
        }
    }
}
