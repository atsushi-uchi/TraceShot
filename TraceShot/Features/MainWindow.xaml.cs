using NHotkey;
using ScreenRecorderLib;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Shell;
using System.Windows.Threading;
using TraceShot.Controls;
using TraceShot.Extensions;
using TraceShot.Models;
using TraceShot.Services;
using TraceShot.ViewModels;
using Windows.Media.SpeechRecognition;
using static TraceShot.Properties.Settings;
using Brushes = System.Windows.Media.Brushes;
using Canvas = System.Windows.Controls.Canvas;
using Cursors = System.Windows.Input.Cursors;
using Drawing = System.Drawing;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using Point = System.Windows.Point;
using Rectangle = System.Windows.Shapes.Rectangle;
using Size = System.Windows.Size;
using TextBox = System.Windows.Controls.TextBox;
using Thumb = System.Windows.Controls.Primitives.Thumb;

namespace TraceShot.Features
{
    // 録画範囲の種類を定義
    public enum RecordMode { FullScreen, Region, Window }

    public partial class MainWindow : Window
    {
        private enum ResizeDirection { None, Left, Right, Top, Bottom, Move }

        public MainViewModel Data { get; } = new();
        private ExportCacheManager _cacheManager = new();
        private SpeechRecognizer? _winrtRecognizer;
        private Dictionary<(ModifierKeys, Key), Action>? _keyBindings;
        private bool _isPlaying = false;
        private bool _isRecording = false;
        private bool _isInternalSelectionChange = false;
        private string _currentVideoPath = "";
        private bool _isDragging = false; // スライダー操作中かどうかの判定
        private bool _isSpeechInitalized = false;
        private DispatcherTimer _playerTimer;
        private MouseHook _mouseHook = new();
        private Drawing.Rectangle? _selectedRegion = null; // 選択された範囲を保持
        private IntPtr _targetWindowHandle;
        string _fullDeviceName = string.Empty;
        string _rectDeviceName = string.Empty;
        private DebugWindow? _debugWindow;

        public MainWindow()
        {
            InitializeComponent();

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Title = $"TraceShot v{version?.Major}.{version?.Minor}.{version?.Build}";

            DataContext = Data;

            Data.ScrollIntoViewRequested = (entry) =>
            {
                App.Current.Dispatcher.Invoke(() =>
                {
                    BookmarkListBox.SelectedItem = entry;
                    BookmarkListBox.ScrollIntoView(entry);
                    BookmarkListBox.UpdateLayout();
                    var container = BookmarkListBox.ItemContainerGenerator.ContainerFromItem(entry) as ListBoxItem;
                    if (container != null)
                    {
                        container.Focus();
                    }
                }, DispatcherPriority.Render);
            };

            Data.RefreshCanvas = () =>
            {
                App.Current.Dispatcher.Invoke(() =>
                {
                    RefreshBookmarkCanvas();
                }, DispatcherPriority.Background);
            };
            Data.GetCurrentPosition = () => VideoPlayer.Position;
            Data.GetVideoSnapshotFunc = () => new VideoSnapshotInfo(VideoPlayer);

            ApplyCurrentSettings();

            if (Data.Config.IsVoiceEnabled)
            {
                InitSpeechRecognition();
            }

            InitializeKeyBindings();

            _playerTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _playerTimer.Tick += Timer_Tick;
            _playerTimer.Start();

            Data.Config.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(Data.CurrentMode))
                {

                    if (Data.CurrentMode == AppViewMode.Edit)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            RefreshBookmarkCanvas();
                        });
                    }
                }
            };
            Data.ShutterRequested += (sender, args) =>
            {
                TriggerShutterEffect();
            };

            var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
            PreviewImage.Source = ImageService.GetReadyStandardImage(dpi);

            // マウスクリックをフックして録画
            Action<int, int> onMouseClick = (x, y) =>
            {
                if (_isRecording) AddClickTriggerBookmark(x, y);
            };
            _mouseHook.OnMouseMiddleClick += onMouseClick;
            _mouseHook.OnSideButton1Click += onMouseClick;
            _mouseHook.OnSideButton2Click += onMouseClick;
            _mouseHook.Start();

            // 録画サービス イベント登録
            RecService.Instance.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;
            RecService.Instance.RecordingErrorOccurred += (s, error) => Data.SwitchToRescueMode(error);
        }

        private void InitializeKeyBindings()
        {
            _keyBindings = new()
            {
                { (ModifierKeys.Control, Key.S), () => SaveEvidence_Click(this, null) },
                { (ModifierKeys.Control, Key.O), () => OpenEvidence_Click(this, null) },
                { (ModifierKeys.Control, Key.E), () => OpenExport_Click(this, null) },
                { (ModifierKeys.Control, Key.C), () => OnCopyMenu_Click(this, null) },
                { (ModifierKeys.Control, Key.V), () => OnPasteMenu_Click(this, null) },
                { (ModifierKeys.Control, Key.Z), () => OnUndo_Click(this, null) },
                { (ModifierKeys.Control, Key.Y), () => OnRedo_Click(this, null) },
                { (ModifierKeys.None, Key.Delete), () => DeleteBookmarkButton_Click(this, null) },
                { (ModifierKeys.None, Key.Escape), () => ClearAnnotationSelection() },
            };
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                TimelineSlider.Maximum = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                string current = VideoPlayer.Position.ToString(@"mm\:ss");
                string total = VideoPlayer.NaturalDuration.TimeSpan.ToString(@"mm\:ss");
                TimeText.Text = $"{current} / {total}";

                bool isUserInteracting = _isDragging || Mouse.LeftButton == MouseButtonState.Pressed;
                if (!isUserInteracting)
                {
                    TimelineSlider.Value = VideoPlayer.Position.TotalSeconds;
                }

                if (_isPlaying)
                {
                    if (StopAtPointCheckBox.IsChecked ?? false)
                    {
                        // 現在のポジション付近にエントリーが存在するか
                        var entry = RecService.Instance.Evidence.Entries
                            .Where(bm => Math.Abs(bm.Time.TotalSeconds - VideoPlayer.Position.TotalSeconds) < 0.1)
                            .OrderBy(bm => bm.Time)
                            .FirstOrDefault();

                        if (entry != null)
                        {
                            if (Data.SelectedItem != entry)
                            {
                                Data.SelectedItem = entry;
                            }
                        }
                        else
                        {
                            Data.SelectedItem = null;
                        }
                    }
                }
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // 現在の入力状態（修飾キー + 押されたキー）をタプルで作成
            var input = (Keyboard.Modifiers, e.Key);

            // 辞書にその組み合わせが登録されていれば実行する
            if (_keyBindings.TryGetValue(input, out var action))
            {
                action();
                e.Handled = true; // イベントが処理されたことを通知（重要！）
            }
        }

        private void ClearAnnotationSelection()
        {
            if (Data.AnnotationManager.SelectedAnnotation != null)
            {
                Data.AnnotationManager.SelectedAnnotation = null;
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            _mouseHook.Stop();

            if (_isRecording)
            {
                RecService.Instance.StopRecording();
            }
        }

        private void ApplyCurrentSettings()
        {
            bool update = false;
            string savedPath = Default.SavePath;
            if (string.IsNullOrEmpty(savedPath) || !Directory.Exists(savedPath))
            {
                // デフォルトの保存先が未設定の場合、MYVideosに設定
                savedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
                Default.SavePath = savedPath;
                update = true;
            }
            SavePathStatusText.Text = savedPath;
            FrameRateText.Text = $"FPS: {Default.FrameRate}";

            if (update)
            {
                Default.Save();
            }
            // ホットキー登録
            RefreshHotkey();

            // マウス中央、サイドボタンの撮影設定
            _mouseHook.ChatteringThreshold = Default.ChatteringThreshold;
            _mouseHook.EnableMiddleClick = Default.EnableMiddleClick;
            _mouseHook.EnableSideClick = Default.EnableSideClick;
        }

        private void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWin = new SettingsWindow();
            settingsWin.Owner = this;
            if (settingsWin.ShowDialog() == true)
            {
                ApplyCurrentSettings();

                if (Data.Config.IsVoiceEnabled && !_isSpeechInitalized)
                {
                    InitSpeechRecognition();
                }
                Data.StatusText = "🛠️ 設定を更新しました";
            }
        }

        private void DebugButton_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = (MenuItem)sender;
            if (menuItem.IsChecked)
            {
                // --- デバッグ ON の処理 ---
                if (_debugWindow == null || !_debugWindow.IsLoaded)
                {
                    _debugWindow = new DebugWindow();
                    _debugWindow.Owner = this;

                    // ウィンドウが閉じられた時の連動
                    _debugWindow.Closed += (s, args) =>
                    {
                        menuItem.IsChecked = false;
                        menuItem.Header = "デバッグ ON"; // 次に押した時のアクションを表示
                        _debugWindow = null;
                    };
                    _debugWindow.Show();
                    menuItem.Header = "デバッグ OFF"; // 現在の状態（または停止操作）を表示
                }
            }
            else
            {
                // --- デバッグ OFF の処理 ---
                if (_debugWindow != null && _debugWindow.IsLoaded)
                {
                    _debugWindow.Close();
                }
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void SaveEvidence_Click(object? sender, RoutedEventArgs? e)
        {
            if (RecService.Instance.Evidence == null) return;

            try
            {
                RecService.Instance.SaveEvidenceJson();

                Data.StatusText = $"[保存完了] {DateTime.Now:HH:mm:ss} エビデンスを保存しました。";
                MessageBox.Show("エビデンスの内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void OpenEvidence_Click(object sender, RoutedEventArgs? e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json",
                Title = "保存されたエビデンス（JSON）を選択してください",
                InitialDirectory = Default.SavePath
            };

            if (openFileDialog.ShowDialog() == true)
            {
                if (await Data.LoadEvidenceAsync(openFileDialog.FileName))
                {
                    var entry = RecService.Instance.Entries.OrderBy(b => b.Time).FirstOrDefault();
                    if (entry != null)
                    {
                        Data.SelectedItem = entry;
                        VideoPlayer.Position = Data.SelectedItem.Time;
                    }
                }
            }
        }

        private void OpenExport_Click(object sender, RoutedEventArgs? e)
        {
            Debug.WriteLine($"OpenExport_Click width={RescueImage.Width} height={RescueImage.Height} ActualWidth={RescueImage.ActualWidth} ActualHeight={RescueImage.ActualHeight}");
            var exportWin = new ExportWindow(_cacheManager)
            {
                Owner = this // 親ウィンドウをセットして中央に表示
            };
            exportWin.ShowDialog();
        }

        private void SavePathStatusText_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 現在のパスを取得
            string folderPath = SavePathStatusText.Text;

            // パスが未設定の場合は MyVideo フォルダをデフォルトにする
            if (string.IsNullOrEmpty(folderPath) || folderPath == "未設定")
            {
                folderPath = Default.SavePath;
            }

            try
            {
                if (Directory.Exists(folderPath))
                {
                    // エクスプローラーで開く
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = folderPath,
                        UseShellExecute = true, // フォルダを関連付けられたアプリ（エクスプローラー）で開く
                        Verb = "open"
                    });

                    Data.StatusText = "📂 フォルダを開きました";
                }
                else
                {
                    Data.StatusText = "❌ 保存先フォルダが見つかりません";
                }
            }
            catch (Exception ex)
            {
                Data.StatusText = "❌ フォルダを開けませんでした";
            }
        }

        private Bookmark CreateBookmarkAtCurrentTime()
        {
            var currentTime = VideoPlayer.Position;
            var lastEntry = Data.TimelineEntries.FirstOrDefault(x => x.Time > currentTime);

            return new Bookmark
            {
                Time = currentTime,
                CaseId = lastEntry?.CaseId ?? 0,
                Result = TestResult.SS,
                Icon = "🖋️",
                Note = "",
            };
        }

        private void DrawingCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not IInputElement canvas || sender is not FrameworkElement frameworkElement) return;

            var currentPos = e.GetPosition(canvas);
            var actualSize = new Size(frameworkElement.ActualWidth, frameworkElement.ActualHeight);

            if (actualSize.Width <= 0 || actualSize.Height <= 0)
            {
                return;
            }

            if (BookmarkListBox.SelectedItem is not Bookmark entry)
            {
                // 救済モード（動画なし）の場合、現在の再生時間は取れないので
                // 必要に応じて処理を分岐（例：新規作成を禁止するか、0秒で作るか）
                if (Data.CurrentMode == AppViewMode.Rescue)
                {
                    Debug.WriteLine("救済モードでブックマークがない場合は、一旦何もしない");
                    return;
                }
                entry = CreateBookmarkAtCurrentTime();
                Data.TimelineEntries.Add(entry);
                Data.SelectedItem = entry;
                Data.UpdateTimelineGroups();
            }

            if (e.ChangedButton != MouseButton.Left)
            {
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                // NoteAnnotation モード
                Data.AnnotationManager.StartDrawing<NoteAnnotation>(entry, currentPos, actualSize);
            }
            else
            {
                // 通常の RectAnnotation モード
                var annotation = Data.AnnotationManager.StartDrawing<RectAnnotation>(entry, currentPos, actualSize);
                if (annotation is RectAnnotation rect)
                {
                    rect.OcrAction = Data.ExecuteOcrAction;
                }
            }
            RefreshBookmarkCanvas();
            ((IInputElement)sender).CaptureMouse();
        }

        private void DrawingCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var currentPos = e.GetPosition(VideoPlayer);
                var actualPos = new Size
                {
                    Width = VideoPlayer.ActualWidth,
                    Height = VideoPlayer.ActualHeight,
                };
                if (Data.CurrentMode == AppViewMode.Rescue) // レスキューモード用
                {
                    currentPos = e.GetPosition(RescueImage);
                    actualPos = new Size(RescueImage.ActualWidth, RescueImage.ActualHeight);
                }

                Data.AnnotationManager.UpdateDrawing(currentPos, actualPos);
            }
        }

        private void DrawingCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (BookmarkListBox.SelectedItem is Bookmark bookmark)
            {
                Data.AnnotationManager.CompleteDrawing(bookmark);

                bookmark.IsDirty = true;
            }

            ((IInputElement)sender).ReleaseMouseCapture();
        }

        private void OnAnnotation_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is AnnotationBase annotation)
            {
                // 1. 基準要素の特定と有効性チェック
                FrameworkElement referenceElement = (Data.CurrentMode == AppViewMode.Rescue) ? RescueImage : VideoPlayer;
                if (!referenceElement.IsVisible) return;

                // 2. 座標取得とテキスト中心オフセット補正
                var transform = thumb.TransformToVisual(referenceElement);
                Point currentPos = transform.Transform(new Point(0, 0));

                var tag = thumb.Tag?.ToString() ?? "";
                if (annotation is NoteAnnotation note && tag == "End")
                {
                    currentPos.X += note.ActualTextWidth / 2.0;
                    currentPos.Y += note.ActualTextHeight / 2.0;
                }

                var actualSize = new Size(referenceElement.ActualWidth, referenceElement.ActualHeight);

                // --- 修正：先に座標を確定させる ---
                // これにより annotation 内部の RelX 等が最新になる
                Data.AnnotationManager.CompleteDrawing(annotation, currentPos, actualSize, tag);

                // 4. アン一ドゥ/リドゥ用の処理（最新の状態を保存）
                if (Data.AnnotationManager.TryConsumeUpdateStart(annotation.Id, out var before))
                {
                    var after = new AnnotationManager.UpdateState
                    {
                        RelX = annotation.RelX,
                        RelY = annotation.RelY,
                        RelWidth = annotation.RelWidth,
                        RelHeight = annotation.RelHeight,
                    };

                    // Note特有のプロパティも忘れずに
                    if (annotation is NoteAnnotation noteAnno)
                    {
                        after.RelStartX = noteAnno.RelStartX;
                        after.RelStartY = noteAnno.RelStartY;
                    }

                    Data.AnnotationManager.PushUpdateAction(annotation, before!, after, Data.SelectedItem);
                }

                // --- 5. ダーティフラグ更新 ---
                if (annotation is CropAnnotation)
                {
                    foreach (var bookmark in Data.Recorder.Entries) { bookmark.Modified(); }

                    Data.AnnotationManager.RefreshCropOverlay();
                }
                else if (BookmarkListBox.SelectedItem is Bookmark bookmark)
                {
                    bookmark.Modified();
                }
            }
        }

        private void OnAnnotation_DragStarted(object sender, DragStartedEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is AnnotationBase annotation)
            {
                Data.AnnotationManager.RecordUpdateStart(annotation);
            }
        }

        private void TextInput_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var textBox = sender as TextBox;
            var note = textBox?.DataContext as NoteAnnotation;
            var bookmark = BookmarkListBox.SelectedItem as Bookmark;

            if (note == null) return;

            switch (e.Key)
            {
                case Key.Enter:
                    // Ctrl や Shift が押されていない「Enter単体」の時だけ確定
                    if (Keyboard.Modifiers == ModifierKeys.None)
                    {
                        CommitNoteEdit(note, bookmark);
                        e.Handled = true;
                    }
                    break;

                case Key.Escape:
                    // ESC キーでキャンセル
                    CancelNoteEdit(note, bookmark);
                    e.Handled = true;
                    break;
            }
        }

        private void TextInput_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            var note = textBox?.DataContext as NoteAnnotation;
            var bookmark = BookmarkListBox.SelectedItem as Bookmark;

            if (note == null) return;

            // フォーカスを失った時は編集を確定
            CommitNoteEdit(note, bookmark);
        }

        private void StartNoteEdit(NoteAnnotation note)
        {
            if (note == null) return;

            note.OriginText = note.Text;
            note.IsEditing = true;
            note.IsCommitted = false;
        }

        private void CommitNoteEdit(NoteAnnotation note, Bookmark? bookmark)
        {
            if (note == null) return;

            if (string.IsNullOrWhiteSpace(note.Text))
            {
                // テキストが空の場合は削除
                Data.AnnotationManager.Remove(bookmark, note);
            }
            else
            {
                if (note.OriginText != note.Text)
                {
                    Data.AnnotationManager.PushTextUpdateAction(note, bookmark);
                }

                // テキストがある場合は確定
                note.IsEditing = false;
                note.IsCommitted = true;
            }
        }

        private void CancelNoteEdit(NoteAnnotation note, Bookmark? bookmark)
        {
            if (note == null) return;

            if (string.IsNullOrEmpty(note.OriginText))
            {
                // 元のテキストがない場合は削除
                Data.AnnotationManager.Remove(bookmark, note);
            }
            else
            {
                // 元のテキストに復元して確定
                note.Text = note.OriginText;
                note.IsEditing = false;
                note.IsCommitted = true;
            }
        }

        private void TextBox_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            // Visible になった（表示された）タイミングかチェック
            if (sender is System.Windows.Controls.Border border && (bool)e.NewValue)
            {
                // Border の中にある子要素（TextBox）を取得
                var textBox = border.Child as TextBox;

                if (textBox != null)
                {
                    // UIスレッドの優先度を少し下げて実行（表示が完了した直後にフォーカスするため）
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        textBox.Focus();
                        // 必要であれば、末尾にカーソルを移動、または全選択
                        textBox.CaretIndex = textBox.Text.Length;
                        // textBox.SelectAll(); 
                    }), System.Windows.Threading.DispatcherPriority.Input);
                }
            }
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                if (sender is System.Windows.Controls.Border border && border.DataContext is NoteAnnotation note)
                {
                    StartNoteEdit(note);
                    e.Handled = true;
                }
            }
        }

        private void OnNoteStart_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is NoteAnnotation note)
            {
                if (Data.CurrentMode == AppViewMode.Edit)
                {
                    note.RelStartX += e.HorizontalChange / VideoPlayer.ActualWidth;
                    note.RelStartY += e.VerticalChange / VideoPlayer.ActualHeight;
                }
                else if (Data.CurrentMode == AppViewMode.Rescue)
                {
                    note.RelStartX += e.HorizontalChange / RescueImage.ActualWidth;
                    note.RelStartY += e.VerticalChange / RescueImage.ActualHeight;
                }
            }
        }

        private void OnNoteText_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is NoteAnnotation note)
                if (!note.IsEditing) // 編集中は移動させない設定
                {
                    if (Data.CurrentMode == AppViewMode.Edit)
                    {
                        // 動画プレイヤーのサイズを基準に、相対的な移動量を加算
                        note.RelX += e.HorizontalChange / VideoPlayer.ActualWidth;
                        note.RelY += e.VerticalChange / VideoPlayer.ActualHeight;
                    }
                    else if(Data.CurrentMode == AppViewMode.Rescue)
                    {
                        note.RelX += e.HorizontalChange / RescueImage.ActualWidth;
                        note.RelY += e.VerticalChange / RescueImage.ActualHeight;
                    }

                    // 念のため、キャンバスからはみ出さないように制限（0.0 〜 1.0）
                    note.RelX = Math.Clamp(note.RelX, 0, 1);
                    note.RelY = Math.Clamp(note.RelY, 0, 1);
                }
        }

        private void OnMoveThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                if (Data.CurrentMode == AppViewMode.Edit)
                {
                    // 変化量(px) / 動画プレイヤーの現在のサイズ(px) = 相対変化量
                    rect.RelX += e.HorizontalChange / VideoPlayer.ActualWidth;
                    rect.RelY += e.VerticalChange / VideoPlayer.ActualHeight;
                }
                else if (Data.CurrentMode == AppViewMode.Rescue)
                {
                    rect.RelX += e.HorizontalChange / RescueImage.ActualWidth;
                    rect.RelY += e.VerticalChange / RescueImage.ActualHeight;
                }

            }
        }

        private void OnResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is not Thumb t || t.DataContext is not RectAnnotation rect) return;

            double deltaX = e.HorizontalChange / VideoPlayer.ActualWidth;
            double deltaY = e.VerticalChange / VideoPlayer.ActualHeight;

            if (Data.CurrentMode == AppViewMode.Rescue) // レスキューモード用
            {
                deltaX = e.HorizontalChange / RescueImage.ActualWidth;
                deltaY = e.VerticalChange / RescueImage.ActualHeight;
            }

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

        private void OnEditMenu_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.DataContext is NoteAnnotation note)
            {
                StartNoteEdit(note);
            }
        }

        private void OnDeleteMenu_Click(object sender, RoutedEventArgs e)
        {
            var bookmark = BookmarkListBox.SelectedItem as Bookmark;

            // MenuItem -> ContextMenu -> 紐付いている Thumb を辿って DataContext を取得
            if (sender is MenuItem menuItem && menuItem.DataContext is AnnotationBase annotation)
            {
                Data.AnnotationManager.Remove(bookmark, annotation);
            }
        }

        private void OnConfirmCrop_Click(object sender, RoutedEventArgs e)
        {
            var existingCrop = Data.AnnotationManager.Annotations.OfType<CropAnnotation>().FirstOrDefault();
            if (existingCrop != null)
            {
                if (RecService.Instance.Evidence.CropState == CropState.Editing)
                {
                    RecService.Instance.Evidence.CropState = CropState.Confirmed;
                }
                else
                {
                    RecService.Instance.Evidence.CropState = CropState.Editing;
                }
            }
            else
            {
                RecService.Instance.Evidence.CropState = CropState.Editing;
                var crop = new CropAnnotation();
                crop.RelWidth = 0.5;
                crop.RelHeight = 0.5;
                crop.RelX = (1.0 - crop.RelWidth) / 2;
                crop.RelY = (1.0 - crop.RelHeight) / 2;
                Data.AnnotationManager.Annotations.Add(crop);
            }

            Data.AnnotationManager.RefreshCropOverlay();
        }

        private void CropEnabledCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            RecService.Instance.Evidence.CropState = CropState.Confirmed;

            foreach(var bookmark in Data.Recorder.Entries)
            {
                bookmark.Modified();
            }

            Data.AnnotationManager.RefreshCropOverlay();
        }

        private void DeleteBookmarkButton_Click(object? sender, RoutedEventArgs? e)
        {
            // 1. 選択されている項目があるかチェック
            if (BookmarkListBox.SelectedItems.Count == 0)
            {
                Data.StatusText = "ℹ️ 削除する項目を選択してください";
                return;
            }

            // 確認メッセージ（任意）
            var result = MessageBox.Show(
                $"{BookmarkListBox.SelectedItems.Count} 件のチェックポイントを削除しますか？",
                "削除の確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                int index = BookmarkListBox.SelectedIndex;

                // 2. 選択された項目を一度別リストにコピーする
                // (列挙中に元のコレクションを変更するとエラーになるため)
                var bookmarks = BookmarkListBox.SelectedItems.Cast<Bookmark>().ToList();
                foreach (var cp in bookmarks)
                {
                    RecService.Instance.Entries.Remove(cp);
                }

                //Debug.WriteLine($"index:{index} count:{TimelineListBox.Items.Count}");

                if (BookmarkListBox.Items.Count > index)
                {
                    BookmarkListBox.SelectedItem = BookmarkListBox.Items[index];
                }
                else if (BookmarkListBox.Items.Count > 0)
                {
                    BookmarkListBox.SelectedItem = BookmarkListBox.Items[0];
                }

                // 4. 保存とステータス更新
                NoteEditBox.Text = "";
                Data.StatusText = $"🗑️ {bookmarks.Count} 件削除しました";
            }
            RefreshBookmarkCanvas();
        }

        private void VideoPlayer_MediaOpened(object sender, RoutedEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var duration = VideoPlayer.NaturalDuration.TimeSpan;

                // スライダーの最大値を動画の長さに合わせる
                TimelineSlider.Maximum = duration.TotalSeconds;
                TimelineSlider.Value = 0;

                // 「00:00 / 総時間」の形式で表示
                TimeText.Text = $"00:00 / {duration:mm\\:ss}";

                RefreshBookmarkCanvas();
            }
        }

        private void VideoPlayer_MediaFailed(object sender, ExceptionRoutedEventArgs e)
        {
            string errorMessage = $"再生に失敗しました: {e.ErrorException.Message}";
            Data.SwitchToRescueMode(errorMessage);
            PreviewImage.Source = null;
        }

        private void SelectRegionButton_Click(object sender, RoutedEventArgs e)
        {
            // 1. MainWindowを非表示にする
            this.Hide();

            // 全画面モード
            if (ModeComboBox.SelectedIndex == 0)
            {
                // 2. SelectionMoniterを作成して表示
                var sm = new SelectionMoniter();
                if (sm.ShowDialog() == true)
                {
                    _fullDeviceName = sm.DeviceName;
                    Data.StatusText = $"{sm.MoniterName}を選択しました ";
                }
            }
            // 矩形選択モード
            if (ModeComboBox.SelectedIndex == 1)
            {
                var selectionRect = new SelectionRect();
                if (selectionRect.ShowDialog() == true)
                {
                    // 従来の矩形選択の結果を保持
                    // 型変換を行って代入
                    var rect = selectionRect.SelectedRegion;
                    _selectedRegion = new(
                        (int)rect.X,
                        (int)rect.Y,
                        (int)rect.Width,
                        (int)rect.Height
                    );
                    _rectDeviceName = selectionRect.TargetDeviceName;
                    Data.StatusText = $"矩形確定: {selectionRect.TargetDeviceName} x:{rect.X} y:{rect.Y} w:{rect.Width} h:{rect.Height}";
                }
            }
            // ウインドウ選択モード
            if (ModeComboBox.SelectedIndex == 2)
            {
                var sw = new SelectionWindow();
                sw.ShowDialog();

                _targetWindowHandle = sw.SelectedHWnd;
                Data.StatusText = $"録画対象：{sw.SelectedTitle}";
            }

            // MainWindowを再表示してアクティブにする
            this.Show();
            this.Activate();
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isDragging)
            {
                VideoPlayer.Position = TimeSpan.FromSeconds(TimelineSlider.Value);
                PlayerPause(true);
                BookmarkListBox.SelectedItem = null;

                // 何も選択されていない場合はキャンバスを空にする
                Data.AnnotationManager.Annotations.Clear();
                Data.AnnotationManager.RefreshCropOverlay();
            }
        }

        private void Slider_DragStarted(object sender, DragStartedEventArgs e)
        {
            _isDragging = true;
            SliderToolTip.Visibility = Visibility.Visible;
            SliderToolTip.IsOpen = true;
        }

        private void Slider_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            _isDragging = false;
            SliderToolTip.IsOpen = false;
            SliderToolTip.Visibility = Visibility.Collapsed;
            AfterSliderChange();
        }

        private void AfterSliderChange()
        {
            double currentValue = TimelineSlider.Value;
            VideoPlayer.Position = TimeSpan.FromSeconds(currentValue);

            if (Data.TimelineEntries == null || !Data.TimelineEntries.Any()) return;

            // 1. しきい値（0.5秒）以内のものを抽出し
            // 2. 現在地との差が一番小さい順に並べ替え
            // 3. その先頭（最も近いもの）を取得する
            var entry = RecService.Instance.Evidence.Entries
                .Where(e => Math.Abs(e.Time.TotalSeconds - currentValue) < 0.05)
                .OrderBy(e => Math.Abs(e.Time.TotalSeconds - currentValue))
                .FirstOrDefault();

            if (entry != null)
            {
                Data.SelectedItem = entry;
            }
            else
            {
                Data.SelectedItem = null;
            }
            PlayerPause(false);
            Data.StatusText = $"Seek: {VideoPlayer.Position}";
        }

        private void Slider_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            // ドラッグ中のみ処理を行う
            if (_isDragging && sender is Slider slider)
            {
                Point mousePos = e.GetPosition(slider);
                double ratio = mousePos.X / slider.ActualWidth;
                double targetSeconds = slider.Maximum * ratio;

                // 範囲制限
                targetSeconds = Math.Max(0, Math.Min(slider.Maximum, targetSeconds));

                TimeSpan time = TimeSpan.FromSeconds(targetSeconds);
                ToolTipText.Text = time.ToString(@"mm\:ss\.fff");

                // つまみの位置に合わせてツールチップを移動
                SliderToolTip.PlacementTarget = slider;
                SliderToolTip.Placement = PlacementMode.Relative;
                SliderToolTip.HorizontalOffset = mousePos.X;
                SliderToolTip.VerticalOffset = -30;
            }
        }
        
        private void Slider_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            AfterSliderChange();
        }

        private void PlayPauseButton_Click(object? sender, RoutedEventArgs? e)
        {
            if (_isPlaying)
            {
                PlayerPause(false);
            }
            else
            {
                // 再生処理
                VideoPlayer.Play();
                PlayPauseIcon.Text = "⏸️";
                PlayPauseText.Text = "一時停止";
                _isPlaying = true;

                Data.StatusText = "▶ 再生中...";
                //StartPlayerTimer();
            }
        }

        private void PlayerPause(bool withReflash)
        {
            if (withReflash) VideoPlayer.Play();

            // 一時停止処理
            VideoPlayer.Pause();
            PlayPauseIcon.Text = "▶️";
            PlayPauseText.Text = "再生";
            _isPlaying = false;

            Data.StatusText = "⏸ 一時停止中";
        }

        private void RefreshHotkey()
        {
            try
            {
                // 証跡追加
                HotkeyRegister.RegisterBookmark("bookmark",
                    (Key)Default.BookmarkHotkeyKey,
                    (ModifierKeys)Default.BookmarkHotkeyMod,
                    OnBookmark);

                // 音声メモ
                HotkeyRegister.RegisterBookmark("voicememo",
                    (Key)Default.VoiceHotkeyKey,
                    (ModifierKeys)Default.VoiceHotkeyMod,
                    OnVoiceMemo);
            }
            catch
            { /* エラー処理 */
            }
        }

        private void OnBookmark(object? sender, HotkeyEventArgs e)
        {
            Data.AddTimelineEntry();

            e.Handled = true;
        }

        private async void OnVoiceMemo(object? sender, HotkeyEventArgs e)
        {
            Data.AddVoiceEntry();

            e.Handled = true;
        }

        private void OnRecordingStarted()
        {
            // フラグ更新
            _isRecording = true;

            VideoPlayer.Stop();

            RecordingIcon.Foreground = Brushes.Black;
            RecordingIcon.Text = "⏹️";
            RecordingText.Text = "録画停止";
            Data.CanAddEntry = true;
        }

        private void OnRecordingStopped()
        {
            // フラグ更新
            _isRecording = false;

            PreviewImage.Source = ImageService.GetReadyStandardImage();

            RefreshBookmarkCanvas();

            Data.StatusText = "保存完了";
            Data.VideoSource = new Uri(_currentVideoPath);

            RecordingIcon.Foreground = Brushes.Red;
            RecordingIcon.Text = "🔴";
            RecordingText.Text = "録画開始";
            Data.CanAddEntry = true;
        }

        private async void RecordingButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRecording)
            {
                try
                {
                    // 1. 保存されている設定を読み込む

                    string modeName = (ModeComboBox.SelectionBoxItem as string) ?? ModeComboBox.Text;
                    _currentVideoPath = RecService.Instance.PrepareEvidence(Default.SavePath, modeName);
                    switch (ModeComboBox.SelectedIndex)
                    {
                        case 0: // 全画面
                            RecService.Instance.StartFullscreenRecording(_currentVideoPath, _fullDeviceName);
                            break;

                        case 1: // 矩形選択
                            if (_selectedRegion == null)
                            {
                                Data.StatusText = "エラー：範囲を先に選択してください";
                                return;
                            }
                            RecService.Instance.StartRectangleRecording(_currentVideoPath, _rectDeviceName, _selectedRegion);
                            break;

                        case 2: // ウィンドウ選択
                            if (_targetWindowHandle == IntPtr.Zero)
                            {
                                Data.StatusText = "エラー：ウィンドウを先に選択してください";
                                return;
                            }
                            RecService.Instance.StartWindowRecording(_currentVideoPath, _targetWindowHandle);
                            break;
                    }

                    OnRecordingStarted();

                    Data.StatusText = $"● 録画中: {RecService.Instance.CurrentVideoName}";

                    taskbarInfo.ProgressState = TaskbarItemProgressState.Error;
                    taskbarInfo.ProgressValue = 1.0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("録画開始エラー: " + ex.Message);
                }
            }
            else
            {
                Data.CanAddEntry = false;

                RecService.Instance.StopRecording();
                Data.StatusText = "動画を処理中...";
                await Task.Delay(1000);

                OnRecordingStopped();
                Data.IsEditMode = true;
            }
        }

        private void AddClickTriggerBookmark(int x, int y)
        {
            TriggerShutterEffect();

            Data.AddTimelineEntry();
        }

        private void RecorderManager_OnPreviewFrameReceived(object? sender, FrameRecordedEventArgs e)
        {
            // UIスレッドで実行
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (RecordingModeGrid.Visibility != Visibility.Visible) return;

                var data = e.BitmapData;
                if (data == null) return;

                int width = data.Width;
                int height = data.Height;

                var bitmap = RecService.Instance.PreviewBitmap;
                if (bitmap == null || bitmap.PixelWidth != data.Width || bitmap.PixelHeight != data.Height)
                {
                    bitmap = new WriteableBitmap(data.Width, data.Height, 96, 96, PixelFormats.Bgr32, null);
                    RecService.Instance.PreviewBitmap = bitmap;
                    PreviewImage.Source = bitmap;
                }
                else if (PreviewImage.Source != bitmap)
                {
                    PreviewImage.Source = bitmap;
                }

                try
                {
                    bitmap.WritePixels(
                        new Int32Rect(0, 0, data.Width, data.Height),
                        data.Data,
                        data.Stride * data.Height,
                        data.Stride);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Preview Update Error: {ex.Message}");
                }
            }), DispatcherPriority.Render);
        }

        private void BookmarkListBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Left || e.Key == Key.Right)
            {
                var listBox = (System.Windows.Controls.ListBox)sender;
                var currentIndex = listBox.SelectedIndex;
                int nextIndex = e.Key == Key.Right ? currentIndex + 1 : currentIndex - 1;

                if (nextIndex >= 0 && nextIndex < listBox.Items.Count)
                {
                    listBox.SelectedIndex = nextIndex;
                    listBox.ScrollIntoView(listBox.SelectedItem);

                    var container = listBox.ItemContainerGenerator.ContainerFromIndex(nextIndex) as UIElement;
                    container?.Focus();

                    e.Handled = true;
                }
            }
            else if (e.Key == Key.Down)
            {
                var listBox = (System.Windows.Controls.ListBox)sender;
                var currentIndex = listBox.SelectedIndex;
                var caseId = (listBox.Items[currentIndex] as Bookmark)?.CaseId ?? 0;

                for (int nextIndex = currentIndex; nextIndex < listBox.Items.Count; nextIndex++)
                {
                    if (listBox.Items[nextIndex] is Bookmark next && next.CaseId != caseId)
                    {
                        listBox.SelectedIndex = nextIndex;
                        listBox.ScrollIntoView(listBox.SelectedItem);

                        var container = listBox.ItemContainerGenerator.ContainerFromIndex(nextIndex) as UIElement;
                        container?.Focus();

                        e.Handled = true;
                        break;
                    }
                }
            }
            else if (e.Key == Key.Up)
            {
                var listBox = (System.Windows.Controls.ListBox)sender;
                var currentIndex = listBox.SelectedIndex;
                if (currentIndex <= 0) return;

                var currentItem = listBox.Items[currentIndex] as Bookmark;
                var currentCaseId = currentItem?.CaseId;
                int targetIndex = -1;

                // A. まず、現在選択しているケースの「さらに上」に同じCaseIdがあるかチェック
                //（＝今、ケースの途中にいるか？）
                if (currentIndex > 0 && (listBox.Items[currentIndex - 1] as Bookmark)?.CaseId == currentCaseId)
                {
                    // 現在のケースの先頭を探す
                    for (int i = currentIndex - 1; i >= 0; i--)
                    {
                        if ((listBox.Items[i] as Bookmark)?.CaseId == currentCaseId)
                            targetIndex = i;
                        else
                            break;
                    }
                }
                else
                {
                    // B. すでにケースの先頭にいるなら、前のケースの先頭を探す（既存ロジック）
                    for (int i = currentIndex - 1; i >= 0; i--)
                    {
                        if (listBox.Items[i] is Bookmark prev && prev.CaseId != currentCaseId)
                        {
                            var prevCaseId = prev.CaseId;
                            targetIndex = i;
                            for (int j = i - 1; j >= 0; j--)
                            {
                                if ((listBox.Items[j] as Bookmark)?.CaseId == prevCaseId)
                                    targetIndex = j;
                                else
                                    break;
                            }
                            break;
                        }
                    }
                }

                if (targetIndex != -1)
                {
                    listBox.SelectedIndex = targetIndex;
                    listBox.ScrollIntoView(listBox.SelectedItem);
                    // コンテナの取得とフォーカス強制
                    listBox.UpdateLayout();
                    (listBox.ItemContainerGenerator.ContainerFromIndex(targetIndex) as UIElement)?.Focus();
                    e.Handled = true;
                }
            }
        }

        private void BookmarkListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BookmarkListBox.SelectedItem is Bookmark selected)
            {
                // --- 救済モード時の画像更新 ---
                if (Data.CurrentMode == AppViewMode.Rescue)
                {
                    if (File.Exists(selected.ImagePath))
                    {
                        var bitmap = Data.LoadImageFromFile(selected.ImagePath);
                        if (bitmap != null)
                        {
                            Data.RescueImageSource = bitmap;

                            // レイアウトを強制更新してCanvasサイズを確定させる
                            RescueModeGrid.UpdateLayout();
                        }
                    }
                }
                else if (Data.CurrentMode == AppViewMode.Edit)
                {
                    if (!_isInternalSelectionChange)
                    {
                        VideoPlayer.Position = selected.Time;
                        PlayerPause(true);
                        Data.StatusText = $"Seek: {selected.Time}";
                    }
                }

                Data.AnnotationManager.LoadAnnotationsFromBookmark(selected);
            }
            else
            {
                // 何も選択されていない場合はキャンバスを空にする
                Data.AnnotationManager.Annotations.Clear();
                Data.AnnotationManager.RefreshCropOverlay();
            }
            RefreshBookmarkCanvas();
        }

        private void RefreshBookmarkCanvas()
        {
            if (RecService.Instance.Evidence == null || !VideoPlayer.NaturalDuration.HasTimeSpan) return;
            BookmarkCanvas.Children.Clear();

            const double ThumbWidth = 17.0;
            double canvasWidth = BookmarkCanvas.ActualWidth;
            if (canvasWidth <= 0) return;

            double effectiveWidth = canvasWidth - ThumbWidth;
            double totalSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            if (totalSeconds <= 0) return;

            foreach (var entry in Data.TimelineEntries)
            {
                if (entry.Result == TestResult.None && string.IsNullOrEmpty(entry.ImagePath)) continue;

                // 時間の割合 (0.0 ～ 1.0)
                double ratio = entry.Time.TotalSeconds / totalSeconds;
                ratio = Math.Clamp(ratio, 0.0, 1.0);

                double xPos = (ratio * effectiveWidth) + (ThumbWidth / 2.0);

                bool isSelected = BookmarkListBox.SelectedItems.Contains(entry);
                double r = isSelected ? 1.2 : 1.1;

                // 三角形（▲）の作成
                Polygon triangle = new()
                {
                    Points = [
                        new Point(0, 0),
                        new Point(-5 * r, 10 * r),
                        new Point(5 * r, 10 * r)
                    ],
                    Fill = isSelected ? Brushes.Orange : Brushes.LightGray,
                    Stroke = isSelected ? Brushes.Red : Brushes.Gray,
                    StrokeThickness = 1,
                    Cursor = Cursors.Hand,
                    Tag = entry
                };
                Canvas.SetZIndex(triangle, isSelected ? 10 : 0);
                triangle.MouseLeftButtonDown += Marker_MouseLeftButtonDown;
                Canvas.SetLeft(triangle, xPos);
                Canvas.SetTop(triangle, 0);

                BookmarkCanvas.Children.Add(triangle);
            }
            RefreshResultRangeCanvas();
        }

        private void RefreshResultRangeCanvas()
        {
            ResultRangeCanvas.Children.Clear();
            if (!VideoPlayer.NaturalDuration.HasTimeSpan) return;

            double totalSec = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            double canvasWidth = ResultRangeCanvas.ActualWidth;
            if (totalSec <= 0 || canvasWidth <= 0) return;

            if (Data.TimelineView is null) return;

            double start = 0;
            double end = 0;
            Bookmark? prevEntry = null;
            foreach (Bookmark entry in Data.TimelineView)
            {
                // グループ名が変わったら新しい矩形を作成
                if (prevEntry != null && prevEntry.GroupName.In(entry.GroupName))
                {
                    end = entry.Time.TotalSeconds;
                }
                else
                {
                    // 直前の矩形があれば閉じる
                    if (end > 0 && prevEntry is not null)
                    {
                        CreateTimelineRect(prevEntry, start, end, totalSec, canvasWidth);
                    }
                    // 次の矩形の開始を設定
                    start = end;
                    end = entry.Time.TotalSeconds;
                }
                prevEntry = entry;
            }

            if (prevEntry != null)
            {
                CreateTimelineRect(prevEntry, start, end, totalSec, canvasWidth);
            }
        }

        private void CreateTimelineRect(Bookmark entry, double start, double end, double totalSec, double canvasWidth)
        {
            double startX = (start / totalSec) * canvasWidth;
            double endX = (end / totalSec) * canvasWidth;
            double width = (endX - startX) - 2;
            //Debug.WriteLine($"RefreshResultRangeCanvas start={start} end={end} width={width}");
            if (width > 0.1)
            {
                Rectangle rect = new()
                {
                    Width = width,
                    Height = ResultRangeCanvas.Height,
                    Fill = GetBrush(entry.Result),
                    Opacity = 0.5,
                    IsHitTestVisible = false,
                    RadiusX = 1,
                    RadiusY = 1,
                };

                Canvas.SetLeft(rect, startX);
                ResultRangeCanvas.Children.Add(rect);

                if (width > 15) // 幅が狭すぎるときは番号を隠す（文字被り防止）
                {
                    TextBlock noText = new()
                    {
                        Text = (entry.CaseId == 0) ? "--" : $"No.{entry.CaseId}",
                        Foreground = Brushes.White,
                        FontSize = 14,
                        FontWeight = FontWeights.Bold,
                        IsHitTestVisible = false // マウス反応を邪魔しない
                    };

                    // 文字のサイズを測定して中央に配置するための準備
                    noText.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                    double textWidth = noText.DesiredSize.Width;
                    double textHeight = noText.DesiredSize.Height;

                    // 帯の中央に配置
                    double textLeft = startX + (width / 2.0) - (textWidth / 2.0);
                    double textTop = (ResultRangeCanvas.Height / 2.0) - (textHeight / 2.0);

                    Canvas.SetLeft(noText, textLeft);
                    Canvas.SetTop(noText, textTop);

                    ResultRangeCanvas.Children.Add(noText);
                }
            }
        }

        private static SolidColorBrush GetBrush(TestResult result)
        {
            return result switch
            {
                TestResult.OK => Brushes.ForestGreen,
                TestResult.NG => Brushes.Crimson,
                TestResult.PEND => Brushes.Orange,
                TestResult.SS => Brushes.DarkGray,
                _ => Brushes.Transparent // 本当に結果がない場合のみ透明
            };
        }

        private void Marker_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is Bookmark entity)
            {
                // 1. 動画をその時間にジャンプ
                VideoPlayer.Position = entity.Time;

                // 2. (任意) リストボックス等の該当項目を選択状態にする
                BookmarkListBox.SelectedItem = entity;
                BookmarkListBox.ScrollIntoView(entity);

                // Slider側のクリックイベントが動かないように「処理済み」とする
                e.Handled = true;
            }
        }

        private async void InitSpeechRecognition()
        {
            try
            {
                // 日本語エンジンで初期化
                _winrtRecognizer = new SpeechRecognizer(new Windows.Globalization.Language("ja-JP"));

                // 文法のコンパイル（必須）
                await _winrtRecognizer.CompileConstraintsAsync();

                _isSpeechInitalized = true;
            }
            catch (Exception)
            {
                MessageBox.Show("WinRT音声認識の初期化に失敗しました。");
            }
        }

        private async Task<string> StartSpeechToText()
        {
            try
            {
                SoundService.Instance.VoiceStartShutter();

                var result = await _winrtRecognizer?.RecognizeAsync();

                if (result.Status == SpeechRecognitionResultStatus.Success)
                {
                    return result.Text;
                }
                else
                {
                    return "(認識失敗)";
                }
            }
            catch (Exception)
            {
                return "(エラー発生)";
            }
            finally
            {
                System.Media.SystemSounds.Asterisk.Play();
            }
        }

        private void ModeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (ModeToggleButton.IsChecked == true)
            {
                Data.CurrentMode = AppViewMode.Edit;
            }
            else
            {
                Data.CurrentMode = AppViewMode.Recording;
            }
        }

        private void SpeedComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (VideoPlayer == null) return;

            if (SpeedComboBox.SelectedItem is ComboBoxItem item && double.TryParse(item.Tag.ToString(), out double speed))
            {
                // 💡 MediaElementの再生速度を変更
                VideoPlayer.SpeedRatio = speed;
            }
        }

        private void VideoPlayer_MediaEnded(object sender, RoutedEventArgs e)
        {
            PlayerPause(false);
            VideoPlayer.Stop();
        }

        private void NoteBorder_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // sender(Border) から DataContext(NoteAnnotation) を取得
            if (sender is FrameworkElement element && element.DataContext is NoteAnnotation note)
            {
                // 実際のサイズを ViewModel のプロパティにセット
                note.ActualTextWidth = e.NewSize.Width;
                note.ActualTextHeight = e.NewSize.Height;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (BookmarkListBox.SelectedItem is Bookmark bookmark)
            {
                bookmark.IsDirty = true;
            }
        }

        private void TimelineSlider_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            RefreshBookmarkCanvas();
        }

        private void CaesNoEditBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (BookmarkListBox.SelectedItems.Count > 1 && sender is TextBox textBox)
            {
                if (textBox.IsFocused && int.TryParse(textBox.Text, out int newCaseId))
                {
                    var selectedList = BookmarkListBox.SelectedItems.Cast<Bookmark>().ToList();
                    foreach (Bookmark entry in selectedList)
                    {
                        if (entry.CaseId != newCaseId)
                        {
                            entry.CaseId = newCaseId;
                        }
                    }
                }
            }
            Data.UpdateTimelineGroups();
            RefreshBookmarkCanvas();
        }

        private void GroupHeader_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var element = sender as FrameworkElement;
            var group = element?.DataContext as CollectionViewGroup;

            if (group == null) return;

            var itemsToSelect = group.Items.Cast<object>().ToList();

            if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                BookmarkListBox.SelectedItems.Clear();
            }

            foreach (var item in itemsToSelect)
            {
                if (!BookmarkListBox.SelectedItems.Contains(item))
                {
                    BookmarkListBox.SelectedItems.Add(item);
                }
            }

            e.Handled = true;

            BookmarkListBox.Focus();
        }

        private void RenumberCases_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkListBox.SelectedItem is Bookmark selected)
            {
                // No.0 自体を選択して「ここから振り直し」をした場合は、No.1から開始させる
                int currentNewId = (selected.CaseId == 0) ? 1 : selected.CaseId;

                var allEntries = Data.TimelineEntries.OrderBy(b => b.Time).ToList();
                int startIndex = allEntries.IndexOf(selected);
                if (startIndex == -1) return;

                int lastOriginalId = selected.CaseId;

                for (int i = startIndex; i < allEntries.Count; i++)
                {
                    var entry = allEntries[i];

                    // ★ポイント：元のIDが0の項目は「未分類」なので、リ番の計算から完全に無視する
                    if (entry.CaseId == 0)
                    {
                        continue;
                    }

                    // 元々のCaseIdが前の項目（0以外）と違うなら、新しい連番をカウントアップ
                    if (lastOriginalId != 0 && entry.CaseId != lastOriginalId)
                    {
                        currentNewId++;
                    }

                    lastOriginalId = entry.CaseId;
                    entry.CaseId = currentNewId;
                }

                Data.UpdateTimelineGroups();
                RefreshBookmarkCanvas();
            }
        }

        private void DeleteEntry_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = BookmarkListBox.SelectedItems.Cast<Bookmark>().ToList();
            foreach (Bookmark entry in selectedItems)
            {
                RecService.Instance.Entries.Remove(entry);
            }

        }

        private void OnCopyMenu_Click(object sender, RoutedEventArgs? e)
        {
            if (sender is MenuItem menuItem && DataContext is MainViewModel vm)
            {
                var annotation = menuItem.DataContext as AnnotationBase;
                Data.CopyAnnotation(annotation);
            }
            else
            {
                Data.CopyAnnotation();
            }
        }

        private void OnPasteMenu_Click(object sender, RoutedEventArgs? e)
        {
            if (BookmarkListBox.SelectedItem is Bookmark bookmark)
            {
                Data.PasteAnnotation(bookmark);
            }
            else
            {
                bookmark = CreateBookmarkAtCurrentTime();
                Data.TimelineEntries.Add(bookmark);
                Data.SelectedItem = bookmark;
                Data.UpdateTimelineGroups();
                Data.PasteAnnotation(bookmark);
            }
        }

        private void OnFocusMenu_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is RectAnnotation rect)
            {
                // Manager経由で状態変更を実行
                var allRects = Data.AnnotationManager.Annotations.OfType<RectAnnotation>();
                Data.AnnotationManager.ExecuteRectStateChange(allRects, rect, "Focus", Data.SelectedItem);
            }
        }

        private void OnMaskingMenu_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is RectAnnotation rect)
            {
                var allRects = Data.AnnotationManager.Annotations.OfType<RectAnnotation>();
                Data.AnnotationManager.ExecuteRectStateChange(allRects, rect, "Masking", Data.SelectedItem);
            }
        }

        private void OnUndo_Click(object sender, RoutedEventArgs? e)
        {
            //Debug.WriteLine("Undo clicked");
            Data.AnnotationManager.Undo();
        }

        private void OnRedo_Click(object sender, RoutedEventArgs? e)
        {
            //Debug.WriteLine("Redo clicked");
            Data.AnnotationManager.Redo();
        }

        public void TriggerShutterEffect()
        {
            // UIスレッドで実行することを保証
            App.Current.Dispatcher.Invoke(() =>
            {
                var animation = new DoubleAnimationUsingKeyFrames();

                // 1. 0.03秒で不透明度 1.0 (完全に真っ白にする)
                // 瞬きよりも速い速度で叩きつけることで「強い閃光」を演出します
                animation.KeyFrames.Add(new DiscreteDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.03))));

                // 2. 0.1秒間、高輝度を維持 (0.8)
                // ここで「溜め」を作ることで、ユーザーの網膜に焼き付けます
                animation.KeyFrames.Add(new LinearDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));

                // 3. 0.4秒かけて、ゆっくりとフェードアウト
                // この「余韻」が安心感に繋がります
                animation.KeyFrames.Add(new SplineDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5)), new KeySpline(0.4, 0, 1, 1)));

                ShutterEffectPanel.BeginAnimation(Border.OpacityProperty, animation);

                SoundService.Instance.PlayShutter();
            });
        }
    }
}