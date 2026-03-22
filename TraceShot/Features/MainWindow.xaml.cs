using NHotkey;
using ScreenRecorderLib;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
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
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Canvas = System.Windows.Controls.Canvas;
using Cursors = System.Windows.Input.Cursors;
using Drawing = System.Drawing;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
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
        public MainViewModel Vm { get; } = new();

        private AnnotationManager _annotationManager;

        private ExportCacheManager _cacheManager = new();

        private SpeechRecognizer? _winrtRecognizer;

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
        private enum ResizeDirection { None, Left, Right, Top, Bottom, Move }
        private DebugWindow? _debugWindow;

        public MainWindow()
        {
            InitializeComponent();

            _annotationManager = new AnnotationManager();
            DataContext = Vm;

            Vm.ScrollIntoViewRequested = (entry) =>
            {
                App.Current.Dispatcher.Invoke(() =>
                {
                    TimelineListBox.ScrollIntoView(entry);
                    TimelineListBox.SelectedItem = entry;
                }, DispatcherPriority.Background);
            };
            Vm.RefreshCanvas = () =>
            {
                App.Current.Dispatcher.Invoke(() =>
                {
                    RefreshBookmarkCanvas();
                }, DispatcherPriority.Background);
            };
            Vm.GetCurrentPosition = () => VideoPlayer.Position;
            Vm.GetVideoSnapshotFunc = () => new VideoSnapshotInfo(VideoPlayer);

            // XAMLのItemsControlのDataContextにマネージャーをセット（またはBindingを設定）
            AnnotationItemsControl.ItemsSource = _annotationManager.Annotations;

            ApplyCurrentSettings();

            if (Vm.Config.IsVoiceEnabled)
            {
                InitSpeechRecognition();
            }

            this.KeyDown += (s, e) => {
                // Ctrl + S で保存を実行する
                if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.S)
                {
                    SaveEvidence_Click(null, null);
                    e.Handled = true; // 他のコントロールにイベントが流れるのを防ぐ
                }
                // Deleteキーで選択項目を削除
                else if (e.Key == Key.Delete)
                {
                    DeleteBookmarkButton_Click(null, null);
                    e.Handled = true;
                }
            };

            _playerTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _playerTimer.Tick += Timer_Tick;
            _playerTimer.Start();

            RecService.Instance.OnRecordingStopped = () =>
            {
                Dispatcher.Invoke(() => Vm.IsEditMode = true);
            };

            Vm.Config.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(Vm.IsEditMode))
                {
                    // 録画モード（IsEditMode == false）に切り替わった場合
                    if (Vm.IsEditMode)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            //RefreshDrawingCanvas();
                            RefreshBookmarkCanvas();
                        });
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            //DrawingCanvas.Children.Clear();
                        });
                    }
                }
            };

            var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
            PreviewImage.Source = ImageService.GetReadyStandardImage(dpi);

            // マウスクリックをフックして録画
            _mouseHook.OnLeftClick += (pos) =>
            {
                // 録画中のみ、証跡追加を実行
                if (_isRecording)
                {
                    AddClickTriggerBookmark(pos);
                }
            };
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
            int fps = Default.FrameRate;
            if (fps == 0) {
                // フレームレートが未設定の場合、30に設定
                fps = 30;
                Default.FrameRate = fps;
                update = true;
            }

            FrameRateText.Text = $"FPS: {fps}";
            RecService.Instance.FrameRate = fps;

            if (update)
            {
                Default.Save();
            }
            // ホットキー登録
            RefreshHotkey();
        }

        private void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWin = new SettingsWindow();
            settingsWin.Owner = this;
            if (settingsWin.ShowDialog() == true)
            {
                ApplyCurrentSettings();

                if (Vm.Config.IsVoiceEnabled && !_isSpeechInitalized)
                {
                    InitSpeechRecognition();
                }
                Vm.StatusText = "🛠️ 設定を更新しました";
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
                    _debugWindow.Closed += (s, args) => {
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

                Vm.StatusText = $"[保存完了] {DateTime.Now:HH:mm:ss} エビデンスを保存しました。";
                MessageBox.Show("エビデンスの内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void OpenEvidence_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json",
                Title = "保存されたエビデンス（JSON）を選択してください",
                InitialDirectory = Default.SavePath
            };

            if (openFileDialog.ShowDialog() == true)
            {
                await Vm.LoadEvidenceAsync(openFileDialog.FileName);
            }
        }

        private void OpenExport_Click(object sender, RoutedEventArgs e)
        {
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

                    Vm.StatusText = "📂 フォルダを開きました";
                }
                else
                {
                    Vm.StatusText = "❌ 保存先フォルダが見つかりません";
                }
            }
            catch (Exception ex)
            {
                Vm.StatusText = "❌ フォルダを開けませんでした";
                RecService.Instance.TraceLogs.Add($"Explorer Error: {ex.Message}");
            }
        }

        private void DrawingCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var currentPos = e.GetPosition(VideoPlayer);
            var actualPos = new Size(VideoPlayer.ActualWidth, VideoPlayer.ActualHeight);
            if (actualPos.Width <= 0 || actualPos.Height <= 0) return;

            if (TimelineListBox.SelectedItem is not TimelineEntry entry)
            {
                int caseId = 0;
                var currentTime = VideoPlayer.Position;
                var lastEntry = Vm.TimelineEntries.FirstOrDefault(x => x.Time > currentTime);
                if (lastEntry != null)
                {
                    caseId = lastEntry.CaseId;
                }
                entry = new TimelineEntry
                {
                    Time = currentTime,
                    CaseId = caseId,
                    Result = TestResult.SS,
                    Icon = "🖋️",
                    Note = "",
                };
                Vm.TimelineEntries.Add(entry);
                Vm.SelectedItem = entry;
                Vm.UpdateTimelineGroups();
            }

            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                // NoteAnnotation モード
                _annotationManager.StartDrawing<NoteAnnotation>(entry, currentPos, actualPos);
            }
            else
            {
                // 通常の RectAnnotation モード
                var annotation = _annotationManager.StartDrawing<RectAnnotation>(entry, currentPos, actualPos);
                if (annotation is RectAnnotation rect)
                {
                    rect.OcrAction = ExecuteOcrOnAnnotation;
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

                _annotationManager.UpdateDrawing(currentPos, actualPos);
            }
        }

        private void DrawingCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (TimelineListBox.SelectedItem is TimelineEntry bookmark)
            {
                _annotationManager.CompleteDrawing(bookmark);

                bookmark.IsDirty = true;
            }

            ((IInputElement)sender).ReleaseMouseCapture();
        }

        private void OnAnnotation_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is AnnotationBase annotation)
            {
                var transform = thumb.TransformToVisual(VideoPlayer);
                Point currentPos = transform.Transform(new Point(0, 0));

                var actualPos = new Size
                {
                    Width = VideoPlayer.ActualWidth,
                    Height = VideoPlayer.ActualHeight,
                };
                var tag = thumb.Tag?.ToString() ?? "";

                _annotationManager.CompleteDrawing(annotation, currentPos, actualPos, tag);

                if (TimelineListBox.SelectedItem is TimelineEntry bookmark)
                {
                    bookmark.IsDirty = true;
                }
            }
        }

        private void TextInput_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var bookmark = TimelineListBox.SelectedItem as TimelineEntry;
            if (e.Key == Key.Enter)
            {
                // Ctrl や Shift が押されていない「Enter単体」の時だけ確定する
                if (Keyboard.Modifiers == ModifierKeys.None)
                {
                    var textBox = sender as TextBox;
                    if (textBox?.DataContext is NoteAnnotation note)
                    {
                        // 1. テキストが空（またはスペースのみ）かチェック
                        if (string.IsNullOrWhiteSpace(note.Text))
                        {
                            // 空ならマネージャー経由で削除
                            _annotationManager.Remove(bookmark, note);
                        }
                        else
                        {
                            // 文字があれば確定処理
                            note.IsEditing = false;
                            note.IsCommitted = true;
                        }
                    }

                    // イベントをここで終了させ、TextBoxに改行を入れさせない
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.Escape)
            {
                var textBox = sender as TextBox;
                if (textBox?.DataContext is NoteAnnotation note)
                {
                    if (string.IsNullOrEmpty(note.OriginText))
                    {
                        // 空ならマネージャー経由で削除
                        _annotationManager.Remove(bookmark, note);
                    }
                    else
                    {
                        note.Text = note.OriginText;
                        note.IsEditing = false;
                        note.IsCommitted = true;
                    }
                }

                // イベントをここで終了させ、TextBoxに改行を入れさせない
                e.Handled = true;
            }
        }

        private void TextInput_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox?.DataContext is NoteAnnotation note)
            {
                // 1. テキストが空（またはスペースのみ）かチェック
                if (string.IsNullOrWhiteSpace(note.Text))
                {
                    if (TimelineListBox.SelectedItem is TimelineEntry bookmark)
                    {
                        _annotationManager.Remove(bookmark, note);
                    }
                }
                else
                {
                    // 文字があれば確定処理
                    note.IsEditing = false;
                    note.IsCommitted = true;
                }
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
            // ダブルクリック（ClickCount == 2）のときだけ処理する
            if (e.ClickCount == 2)
            {
                var border = sender as System.Windows.Controls.Border;
                var note = border?.DataContext as NoteAnnotation;

                if (note != null)
                {
                    // 編集モードをTrueにする
                    note.OriginText = note.Text;
                    note.IsEditing = true;
                    note.IsCommitted = false;

                    // イベントをここで完了させ、親の Canvas などにクリックが伝わらないようにする
                    e.Handled = true;
                }
            }
        }

        private void OnNoteStart_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is NoteAnnotation note)
            {
                note.RelStartX += e.HorizontalChange / VideoPlayer.ActualWidth;
                note.RelStartY += e.VerticalChange / VideoPlayer.ActualHeight;
            }
        }

        private void OnNoteText_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is NoteAnnotation note)
            if (!note.IsEditing) // 編集中は移動させない設定
            {
                // 動画プレイヤーのサイズを基準に、相対的な移動量を加算
                note.RelX += e.HorizontalChange / VideoPlayer.ActualWidth;
                note.RelY += e.VerticalChange / VideoPlayer.ActualHeight;

                // 念のため、キャンバスからはみ出さないように制限（0.0 〜 1.0）
                note.RelX = Math.Clamp(note.RelX, 0, 1);
                note.RelY = Math.Clamp(note.RelY, 0, 1);
            }
        }

        // 位置を移動させる
        private void OnMoveThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                // 変化量(px) / 動画プレイヤーの現在のサイズ(px) = 相対変化量
                rect.RelX += e.HorizontalChange / VideoPlayer.ActualWidth;
                rect.RelY += e.VerticalChange / VideoPlayer.ActualHeight;
            }
        }


        // 左辺リサイズ：幅を減らした分だけ、Xを右に動かす（逆も然り）
        private void OnLeftResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaRel = e.HorizontalChange / VideoPlayer.ActualWidth;
                rect.RelX += deltaRel;
                rect.RelWidth -= deltaRel;
            }
        }

        // 上辺リサイズ：高さを減らした分だけ、Yを下へ動かす
        private void OnTopResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaRel = e.VerticalChange / VideoPlayer.ActualHeight;
                rect.RelY += deltaRel;
                rect.RelHeight -= deltaRel;
            }
        }

        // 右辺リサイズ：幅をそのまま増減させる
        private void OnRightResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaRel = e.HorizontalChange / VideoPlayer.ActualWidth;
                rect.RelWidth += deltaRel;
            }
        }

        // 下辺リサイズ：高さをそのまま増減させる
        private void OnBottomResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaRel = e.VerticalChange / VideoPlayer.ActualHeight;
                rect.RelHeight += deltaRel;
            }
        }

        // 右下（一番シンプル）：Width と Height をそのまま増減
        private void OnBottomRightResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                rect.RelWidth += e.HorizontalChange / VideoPlayer.ActualWidth;
                rect.RelHeight += e.VerticalChange / VideoPlayer.ActualHeight;
            }
        }

        // 左上（一番複雑）：Width/Height を増やしつつ、X/Y を減らす
        private void OnTopLeftResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaX = e.HorizontalChange / VideoPlayer.ActualWidth;
                double deltaY = e.VerticalChange / VideoPlayer.ActualHeight;

                rect.RelX += deltaX;
                rect.RelWidth -= deltaX;

                rect.RelY += deltaY;
                rect.RelHeight -= deltaY;
            }
        }

        // 右上：Width を増やし、Height を増やしつつ Y を減らす
        private void OnTopRightResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaX = e.HorizontalChange / VideoPlayer.ActualWidth;
                double deltaY = e.VerticalChange / VideoPlayer.ActualHeight;

                rect.RelWidth += deltaX;

                rect.RelY += deltaY;
                rect.RelHeight -= deltaY;
            }
        }

        // 左下：Width を増やしつつ X を減らし、Height を増やす
        private void OnBottomLeftResize_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb t && t.DataContext is RectAnnotation rect)
            {
                double deltaX = e.HorizontalChange / VideoPlayer.ActualWidth;
                double deltaY = e.VerticalChange / VideoPlayer.ActualHeight;

                rect.RelX += deltaX;
                rect.RelWidth -= deltaX;
                rect.RelHeight += deltaY;
            }
        }

        private void OnEditMenu_Click(object sender, RoutedEventArgs e)
        {
            var bookmark = TimelineListBox.SelectedItem as TimelineEntry;

            // MenuItem -> ContextMenu -> 紐付いている Thumb を辿って DataContext を取得
            if (sender is MenuItem menuItem && menuItem.DataContext is NoteAnnotation note)
            {
                
                note.OriginText = note.Text;
                note.IsEditing = true;
                note.IsCommitted = false;
            }
        }

        private void OnDeleteMenu_Click(object sender, RoutedEventArgs e)
        {
            var bookmark = TimelineListBox.SelectedItem as TimelineEntry;

            // MenuItem -> ContextMenu -> 紐付いている Thumb を辿って DataContext を取得
            if (sender is MenuItem menuItem && menuItem.DataContext is AnnotationBase annotation)
            {
                _annotationManager.Remove(bookmark, annotation);
            }
        }

        private void OnConfirmCrop_Click(object sender, RoutedEventArgs e)
        {
            var existingCrop = _annotationManager.Annotations.OfType<CropAnnotation>().FirstOrDefault();
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
                //crop.State = CropState.Editing;
                crop.RelWidth = 0.5;
                crop.RelHeight = 0.5;
                // 中央に配置するための座標計算
                // (1.0 - 0.5) / 2 = 0.25
                crop.RelX = (1.0 - crop.RelWidth) / 2;
                crop.RelY = (1.0 - crop.RelHeight) / 2;
                _annotationManager.Annotations.Add(crop);
            }
            _annotationManager.RefreshCropOverlay();
        }

        private void CropEnabledCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            RecService.Instance.Evidence.CropState = CropState.Confirmed;
            _annotationManager.RefreshCropOverlay();
        }

        private void DeleteBookmarkButton_Click(object? sender, RoutedEventArgs? e)
        {
            // 1. 選択されている項目があるかチェック
            if (TimelineListBox.SelectedItems.Count == 0)
            {
                Vm.StatusText = "ℹ️ 削除する項目を選択してください";
                return;
            }

            // 確認メッセージ（任意）
            var result = MessageBox.Show(
                $"{TimelineListBox.SelectedItems.Count} 件のチェックポイントを削除しますか？",
                "削除の確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                int index = TimelineListBox.SelectedIndex;

                // 2. 選択された項目を一度別リストにコピーする
                // (列挙中に元のコレクションを変更するとエラーになるため)
                var bookmarks = TimelineListBox.SelectedItems.Cast<TimelineEntry>().ToList();
                foreach (var cp in bookmarks)
                {
                    // 3. データソースから削除
                    RecService.Instance.Entries.Remove(cp);
                }

                Debug.WriteLine($"index:{index} count:{TimelineListBox.Items.Count}");

                if (TimelineListBox.Items.Count > index)
                {
                    TimelineListBox.SelectedItem = TimelineListBox.Items[index];
                }
                else if (TimelineListBox.Items.Count > 0)
                {
                    TimelineListBox.SelectedItem = TimelineListBox.Items[0];
                }

                // 4. 保存とステータス更新
                NoteEditBox.Text = "";
                Vm.StatusText = $"🗑️ {bookmarks.Count} 件削除しました";
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
            // 1. ステータスバーにエラーを表示
            Vm.StatusText = "❌ 再生エラー";

            // 2. ログに詳細を記録（以前の画像で見られた TraceLogs や TimelineListBox を活用）
            string errorMessage = $"再生に失敗しました: {e.ErrorException.Message}";
            RecService.Instance.TraceLogs.Add(errorMessage);

            // 3. ユーザーへの通知
            MessageBox.Show(errorMessage, "再生エラー", MessageBoxButton.OK, MessageBoxImage.Error);

            // 必要に応じて代替のプレビュー画像などを表示
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
                    Vm.StatusText = $"{sm.MoniterName}を選択しました ";
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
                    _selectedRegion = new (
                        (int)rect.X,
                        (int)rect.Y,
                        (int)rect.Width,
                        (int)rect.Height
                    );
                    _rectDeviceName = selectionRect.TargetDeviceName;
                    Vm.StatusText = $"矩形確定: {selectionRect.TargetDeviceName} x:{rect.X} y:{rect.Y} w:{rect.Width} h:{rect.Height}";
                }
            }
            // ウインドウ選択モード
            if (ModeComboBox.SelectedIndex == 2)
            {
                var sw = new SelectionWindow();
                sw.ShowDialog();

                _targetWindowHandle = sw.SelectedHWnd;
                Vm.StatusText = $"録画対象：{sw.SelectedTitle}";
            }

            // MainWindowを再表示してアクティブにする
            this.Show();
            this.Activate();
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
            }
        }

        // スライダーを手動で動かしたとき
        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isDragging)
            {
                VideoPlayer.Position = TimeSpan.FromSeconds(TimelineSlider.Value);
                PlayerPause(true);
                TimelineListBox.SelectedItem = null;

                // 何も選択されていない場合はキャンバスを空にする
                _annotationManager.Annotations.Clear();
                _annotationManager.RefreshCropOverlay();
            }
        }

        // スライダーを掴んだらタイマーを止める（操作しやすくするため）
        private void Slider_DragStarted(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e)
        {
            _isDragging = true;
            SliderToolTip.Visibility = Visibility.Visible;
            SliderToolTip.IsOpen = true;
        }

        // 離したら動画の再生位置を確定してタイマー再開
        private void Slider_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            _isDragging = false;
            SliderToolTip.IsOpen = false;
            SliderToolTip.Visibility = Visibility.Collapsed;

            double currentValue = TimelineSlider.Value;
            VideoPlayer.Position = TimeSpan.FromSeconds(currentValue);

            if (Vm.TimelineEntries == null || !Vm.TimelineEntries.Any()) return;

            // 1. しきい値（0.5秒）以内のものを抽出し
            // 2. 現在地との差が一番小さい順に並べ替え
            // 3. その先頭（最も近いもの）を取得する
            var nearbyBookmark = RecService.Instance.Evidence.Entries
                .Where(bm => Math.Abs(bm.Time.TotalSeconds - currentValue) < 0.05)
                .OrderBy(bm => Math.Abs(bm.Time.TotalSeconds - currentValue))
                .FirstOrDefault();

            if (nearbyBookmark != null)
            {
                if (TimelineListBox.SelectedItem != nearbyBookmark)
                {
                    TimelineListBox.SelectedItem = nearbyBookmark;
                    TimelineListBox.ScrollIntoView(nearbyBookmark);
                }
            }
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
                SliderToolTip.Placement = System.Windows.Controls.Primitives.PlacementMode.Relative;
                SliderToolTip.HorizontalOffset = mousePos.X;
                SliderToolTip.VerticalOffset = -30;
            }
        }
        private void Slider_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var slider = sender as Slider;
            if (slider != null)
            {
                var pos = TimeSpan.FromSeconds(slider.Value);
                VideoPlayer.Position = pos;
                PlayerPause(false);
                Vm.StatusText = $"Seek: {pos}";
            }
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

                Vm.StatusText = "▶ 再生中...";
                //StartPlayerTimer();
            }
        }

        public void PlayerPause(bool withReflash)
        {
            if (withReflash) VideoPlayer.Play();

            // 一時停止処理
            VideoPlayer.Pause();
            PlayPauseIcon.Text = "▶️";
            PlayPauseText.Text = "再生";
            _isPlaying = false;

            Vm.StatusText = "⏸ 一時停止中";
        }

        // 録画開始の処理の中に追記
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
            catch { /* エラー処理 */
        }
        }

        // ホットキーが押された時の動作
        private void OnBookmark(object? sender, HotkeyEventArgs e)
        {
            if (RecService.Instance.IsRecording)
            {
                AddBookmarkWhileRecording_Click(this, null);
            }
            else
            {
                AddBookmarkWhilePlaying_Click(this, null);
            }

            e.Handled = true;
        }

        private async void OnVoiceMemo(object? sender, HotkeyEventArgs e)
        {
            if (RecService.Instance.IsRecording)
            {
                AddVoiceMemoRecording_Click(this, null);
            }
            else
            {
                AddVoiceMemoWhilePlaying_Click(this, null);
            }

            e.Handled = true;
        }

        // 録画開始時の処理
        private void OnRecordingStarted()
        {
            // フラグ更新
            _isRecording = true;

            // イベント登録
            RecService.Instance.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

            VideoPlayer.Stop();

            RecordingIcon.Foreground = Brushes.Black;
            RecordingIcon.Text = "⏹️";
            RecordingText.Text = "録画停止";
        }

        // 録画停止時の処理
        private void OnRecordingStopped()
        {
            // フラグ更新
            _isRecording = false;

            // イベント解除
            RecService.Instance.OnPreviewFrameReceived -= RecorderManager_OnPreviewFrameReceived;
            Vm.PreviewBitmap = null;

            PreviewImage.Source = ImageService.GetReadyStandardImage();

            RefreshBookmarkCanvas();

            Vm.StatusText = "保存完了";

            VideoPlayer.Source = new Uri(_currentVideoPath);
            //PlayerPause(true);

            RecordingIcon.Foreground = Brushes.Red;
            RecordingIcon.Text = "🔴";
            RecordingText.Text = "録画開始";
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
                                Vm.StatusText = "エラー：範囲を先に選択してください";
                                return;
                            }
                            RecService.Instance.StartRectangleRecording(_currentVideoPath, _rectDeviceName, _selectedRegion);
                            break;

                        case 2: // ウィンドウ選択
                            if (_targetWindowHandle == IntPtr.Zero)
                            {
                                Vm.StatusText = "エラー：ウィンドウを先に選択してください";
                                return;
                            }
                            RecService.Instance.StartWindowRecording(_currentVideoPath, _targetWindowHandle);
                            break;
                    }

                    OnRecordingStarted();

                    Vm.StatusText = $"● 録画中: {RecService.Instance.CurrentVideoName}";

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
                RecService.Instance.StopRecording();
                Vm.StatusText = "動画を処理中...";
                await Task.Delay(1000);

                OnRecordingStopped();

                if (RecService.Instance.TraceLogs.Count > 0)
                {
                    taskbarInfo.ProgressState = TaskbarItemProgressState.None;
                    taskbarInfo.ProgressValue = 0;
                }
            }
        }

        private void AddClickTriggerBookmark(Drawing.Point rawMousePos)
        {
            SoundService.Instance.PlayShutter();

            if (Vm.PreviewBitmap is not null)
            {
                TimelineEntry newBookmark = new()
                {
                    Time = RecService.Instance.CurrentDuration,
                    Icon = "🖱️",
                    Note = "Click",
                };

                var path = RecService.Instance.SaveBitmap(newBookmark, Vm.PreviewBitmap);
                newBookmark.ImagePath = path;
                Vm.StatusText = $"記録 {newBookmark.Time} {newBookmark.Note} SS作成 {path}";

                RecService.Instance.Entries.Add(newBookmark);
                RefreshBookmarkCanvas();
            }
        }

        private void RecorderManager_OnPreviewFrameReceived(object? sender, FrameRecordedEventArgs e)
        {
            // UIスレッドで実行
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (RecordingOverlay.Visibility != Visibility.Visible) return;

                var data = e.BitmapData;
                if (data == null) return;

                int width = data.Width;
                int height = data.Height;

                // ビットマップの初期化/再作成
                if (Vm.PreviewBitmap == null || Vm.PreviewBitmap.PixelWidth != width || Vm.PreviewBitmap.PixelHeight != height)
                {
                    // ScreenRecorderLibのBitmapDataは通常 Bgr32 (24bitの場合は Bgr24)
                    // アルファチャネルを含む場合は Pbgra32 など調整が必要な場合があります
                    Vm.PreviewBitmap = new WriteableBitmap(width, height, 96, 96, PixelFormats.Bgr32, null);
                    PreviewImage.Source = Vm.PreviewBitmap;
                }

                Vm.PreviewBitmap.Lock();
                try
                {
                    // BitmapData.Scan0 (IntPtr) から WriteableBitmap へコピー
                    // 第3引数の bufferSize は Stride * Height で計算
                    int bufferSize = data.Stride * height;

                    Vm.PreviewBitmap.WritePixels(
                        new Int32Rect(0, 0, width, height),
                        data.Data,
                        bufferSize,
                        data.Stride);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Preview Update Error: {ex.Message}");
                }
                finally
                {
                    Vm.PreviewBitmap.Unlock();
                }
            }));
        }

        private void TimelineListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TimelineListBox.SelectedItem is TimelineEntry selected)
            {
                if (!_isInternalSelectionChange)
                {
                    VideoPlayer.Position = selected.Time;
                    PlayerPause(true);
                    Vm.StatusText = $"Seek: {selected.Time}";
                }
                // マネージャーの表示リストを切り替える
                _annotationManager.LoadAnnotationsFromBookmark(selected);
            }
            else
            {
                // 何も選択されていない場合はキャンバスを空にする
                _annotationManager.Annotations.Clear();
                _annotationManager.RefreshCropOverlay();
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

            foreach (var entry in Vm.TimelineEntries)
            {
                if (entry.Result == TestResult.None && string.IsNullOrEmpty(entry.ImagePath)) continue;

                // 時間の割合 (0.0 ～ 1.0)
                double ratio = entry.Time.TotalSeconds / totalSeconds;
                ratio = Math.Clamp(ratio, 0.0, 1.0);

                double xPos = (ratio * effectiveWidth) + (ThumbWidth / 2.0);

                bool isSelected = TimelineListBox.SelectedItems.Contains(entry);
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

            if (Vm.TimelineView is null) return;

            double start = 0;
            double end = 0;
            TimelineEntry? prevEntry = null;
            foreach (TimelineEntry entry in Vm.TimelineView)
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

        private void CreateTimelineRect(TimelineEntry entry, double start, double end, double totalSec, double canvasWidth)
        {
            double startX = (start / totalSec) * canvasWidth;
            double endX = (end / totalSec) * canvasWidth;
            double width = (endX - startX) - 2;
            Debug.WriteLine($"RefreshResultRangeCanvas start={start} end={end} width={width}");
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
                        Text = (entry.CaseId == 0) ? "SS" : $"No.{entry.CaseId}",
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

        private Brush GetBrush(TestResult result)
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
            if (sender is FrameworkElement el && el.Tag is TimelineEntry entity)
            {
                // 1. 動画をその時間にジャンプ
                VideoPlayer.Position = entity.Time;

                // 2. (任意) リストボックス等の該当項目を選択状態にする
                TimelineListBox.SelectedItem = entity;
                TimelineListBox.ScrollIntoView(entity);

                // Slider側のクリックイベントが動かないように「処理済み」とする
                e.Handled = true;
            }
        }

        // 音声認識の初期化（コンストラクタなどで呼ぶ）
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

        private void AddBookmarkWhileRecording_Click(object sender, RoutedEventArgs? e)
        {
            SoundService.Instance.PlayShutter();

            TimelineEntry bookmark = new()
            {
                Time = RecService.Instance.CurrentDuration,
                Icon = "📌",
                Note = "Add",
            };

            if (Vm.PreviewBitmap is not null)
            {
                var path = RecService.Instance.SaveBitmap(bookmark, Vm.PreviewBitmap);
                bookmark.ImagePath = path;
                Vm.StatusText = $"記録 {bookmark.Time} {bookmark.Note} SS作成 {path}";
            }
            else
            {
                Vm.StatusText = $"記録 {bookmark.Time} {bookmark.Note}";
            }
            RecService.Instance.Entries.Add(bookmark);
            TimelineListBox.SelectedItem = bookmark;
            TimelineListBox.ScrollIntoView(bookmark);
            RefreshBookmarkCanvas();
        }

        private void AddVoiceMemoRecording_Click(object sender, RoutedEventArgs? e)
        {
            TimelineEntry bookmark = new()
            {
                Time = RecService.Instance.CurrentDuration,
                Icon = "🎙",
                Note = "音声入力中",
            };

            RecService.Instance.Entries.Add(bookmark);
            TimelineListBox.SelectedItem = bookmark;
            TimelineListBox.ScrollIntoView(bookmark);
            RefreshBookmarkCanvas();

            Task.Run(async () => {
                try
                {
                    bookmark.IsListening = true;
                    var text = await StartSpeechToText();
                    bookmark.Note = text;
                }
                finally
                {
                    bookmark.IsListening = false;
                }
            });

            if (Vm.PreviewBitmap is not null)
            {
                var path = RecService.Instance.SaveBitmap(bookmark, Vm.PreviewBitmap);
                bookmark.ImagePath = path;
                Vm.StatusText = $"記録 {bookmark.Time} {bookmark.Note} SS作成 {path}";
            }
            else
            {
                Vm.StatusText = $"記録 {bookmark.Time} {bookmark.Note}";
            }
        }

        private void AddBookmarkWhilePlaying_Click(object sender, RoutedEventArgs? e)
        {
            // 選択されているブックマークがある場合は何もしない（誤操作防止）
            if (TimelineListBox.SelectedItem is TimelineEntry selected) return;

            TimelineEntry bookmark = new()
            {
                Time = VideoPlayer.Position,
                Note = "Add",
                Icon = "📋"
            };

            RecService.Instance.Entries.Add(bookmark);
            TimelineListBox.SelectedItem = bookmark;
            TimelineListBox.ScrollIntoView(bookmark);
            RefreshBookmarkCanvas();
        }

        private void AddVoiceMemoWhilePlaying_Click(object sender, RoutedEventArgs? e)
        {
            var bookmark = TimelineListBox.SelectedItem as TimelineEntry;
            if (bookmark is null)
            {
                bookmark = new()
                {
                    Time = VideoPlayer.Position,
                    Note = "音声入力中",
                    Icon = "🎤"
                };
                RecService.Instance.Entries.Add(bookmark);
                TimelineListBox.SelectedItem = bookmark;
                TimelineListBox.ScrollIntoView(bookmark);
            }

            RefreshBookmarkCanvas();

            Task.Run(async () => {
                try
                {
                    bookmark.IsListening = true;
                    var text = await StartSpeechToText();
                    if (string.IsNullOrEmpty(bookmark.Note) || bookmark.Note.Equals("音声入力中"))
                    {
                        bookmark.Note = text;
                    }
                    else
                    {
                        bookmark.Note += $"{Environment.NewLine}{text}";
                    }
                }
                finally
                {
                    bookmark.IsListening = false;
                }
            });
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
                Vm.IsEditMode = true;
                //RefreshDrawingCanvas();
            }
            else
            {
                Vm.IsEditMode = false;
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

        private void ClickTriggerCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _mouseHook.Start();
        }

        private void ClickTriggerCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            _mouseHook.Stop();
        }

        private async Task ExecuteOcrOnAnnotation(RectAnnotation rect)
        {
            var bm = TimelineListBox.SelectedItem as TimelineEntry;
            if (bm == null) return;

            // 1. ビデオ情報を取得
            var info = new VideoSnapshotInfo(VideoPlayer);

            // 2. 正味の画像を生成（既存の ImageService を利用）
            BitmapSource pureVideoBitmap = ImageService.GeneratePureVideoBitmap(bm, info);

            // 3. RectAnnotation の相対座標（RelX, RelY...）を使用
            // 0.0〜1.0 なので、そのまま PixelWidth/Height を掛けるだけ
            int px = (int)(rect.RelX * pureVideoBitmap.PixelWidth);
            int py = (int)(rect.RelY * pureVideoBitmap.PixelHeight);
            int pw = (int)(rect.RelWidth * pureVideoBitmap.PixelWidth);
            int ph = (int)(rect.RelHeight * pureVideoBitmap.PixelHeight);

            // 範囲チェック（画像の外にはみ出さないように）
            px = Math.Clamp(px, 0, pureVideoBitmap.PixelWidth - 1);
            py = Math.Clamp(py, 0, pureVideoBitmap.PixelHeight - 1);
            pw = Math.Min(pw, pureVideoBitmap.PixelWidth - px);
            ph = Math.Min(ph, pureVideoBitmap.PixelHeight - py);

            if (pw <= 0 || ph <= 0) return;

            var cropped = new CroppedBitmap(pureVideoBitmap, new Int32Rect(px, py, pw, ph));

            // 4. OCR実行
            Vm.StatusText = "解析中...";
            string result = await ImageService.RecognizeTextFromBitmapSource(cropped);

            if (!string.IsNullOrWhiteSpace(result))
            {
                string cleanText = result.Replace("\r", "").Replace("\n", " ").Trim();
                bm.AddNewLine(cleanText); // Bookmarkにテキストを追加
                Vm.StatusText = "OCR完了";
            }
            else
            {
                Vm.StatusText = "文字を検出できませんでした";
            }
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
            if (TimelineListBox.SelectedItem is TimelineEntry bookmark)
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
            if (TimelineListBox.SelectedItems.Count > 1 && sender is TextBox textBox)
            {
                if (textBox.IsFocused && int.TryParse(textBox.Text, out int newCaseId))
                {
                    var selectedList = TimelineListBox.SelectedItems.Cast<TimelineEntry>().ToList();
                    foreach (TimelineEntry entry in selectedList)
                    {
                        if (entry.CaseId != newCaseId)
                        {
                            entry.CaseId = newCaseId;
                        }
                    }
                }
            }
            Vm.UpdateTimelineGroups();
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
                TimelineListBox.SelectedItems.Clear();
            }

            foreach (var item in itemsToSelect)
            {
                if (!TimelineListBox.SelectedItems.Contains(item))
                {
                    TimelineListBox.SelectedItems.Add(item);
                }
            }

            e.Handled = true;

            TimelineListBox.Focus();
        }
    }
}