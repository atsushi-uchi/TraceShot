using NHotkey;
using ScreenRecorderLib;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Shell;
using System.Windows.Threading;
using TraceShot.Models;
using TraceShot.Services;
using Windows.Media.SpeechRecognition;
using static TraceShot.Properties.Settings;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Canvas = System.Windows.Controls.Canvas;
using Color = System.Windows.Media.Color;
using Cursors = System.Windows.Input.Cursors;
using Drawing = System.Drawing;
using Line = System.Windows.Shapes.Line;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Point = System.Windows.Point;
using TextBox = System.Windows.Controls.TextBox;
using WpfPoint = System.Windows.Point; // WPFの座標
using WpfRectangle = System.Windows.Shapes.Rectangle;

namespace TraceShot.Features
{
    // 録画範囲の種類を定義
    public enum RecordMode { FullScreen, Region, Window }


    public partial class MainWindow : Window
    {
        private bool _isRecordMode = true;

        public bool IsRecordMode
        {
            get => _isRecordMode;
            set
            {
                _isRecordMode = value;
                UpdateModeUI(); // 💡 値が変わったらUIを強制更新する
            }
        }

        // 💡 モードに合わせてコントロールの表示・非表示を一括制御
        private void UpdateModeUI()
        {
            if (_isRecordMode)
            {
                VideoPlayer.Visibility = Visibility.Collapsed;
                PlayerPanel.Visibility = Visibility.Collapsed;

                RecordingOverlay.Visibility = Visibility.Visible;
                RecordingTimerArea.Visibility = Visibility.Visible;
                RecordingPanel.Visibility = Visibility.Visible;
                RecordModeMenuItem.IsChecked = true;
                PlayerModeMenuItem.IsChecked = false;
            }
            else
            {
                VideoPlayer.Visibility = Visibility.Visible;
                PlayerPanel.Visibility = Visibility.Visible;

                RecordingOverlay.Visibility = Visibility.Collapsed;
                RecordingTimerArea.Visibility = Visibility.Collapsed;
                RecordingPanel.Visibility = Visibility.Collapsed;
                RecordModeMenuItem.IsChecked = false;
                PlayerModeMenuItem.IsChecked = true;
            }
        }

        private SpeechRecognizer? _winrtRecognizer;
        private SettingsService _setting = SettingsService.Instance;
        private bool _isPlaying = false;
        private bool _isRecording = false;
        private bool _isInternalSelectionChange = false;
        private string _currentVideoPath = "";
        private DispatcherTimer _playerTimer;
        private bool _isDragging = false; // スライダー操作中かどうかの判定

        private Drawing.Rectangle? _selectedRegion = null; // 選択された範囲を保持

        private IntPtr _targetWindowHandle;
        private WpfPoint _startPoint;
        private WpfPoint _endPoint;
        
        private WpfRectangle? _currentRectangle;
        private WriteableBitmap? _previewBitmap;
        string _fullDeviceName = string.Empty;
        string _rectDeviceName = string.Empty;
        private Line? _dragLine;
        private bool _isDrawing = false; // 💡 描画中かどうかを管理
        private bool _isResizing = false; // 追加
        private FrameworkElement ?_draggingRect = null;
        private WpfPoint _lastMousePosition; // 直前のマウス座標
        private TextBox? _activeBalloonInput;
        private enum ResizeDirection { None, Left, Right, Top, Bottom, Move }
        private ResizeDirection _currentResizeDir = ResizeDirection.None;

        public MainWindow()
        {
            InitializeComponent();

            DataContext = _setting;

            ApplyCurrentSettings();

            if (_setting.IsVoiceEnabled)
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

            _playerTimer = new DispatcherTimer();
            _playerTimer.Interval = TimeSpan.FromMilliseconds(200); // 0.2秒ごとに更新
            _playerTimer.Tick += Timer_Tick;

            // 1. 最初の一回だけ紐付けを行う
            BookmarkListBox.ItemsSource = RecManager.Instance.Bookmarks;

            // 2. ソートの初期設定
            var view = CollectionViewSource.GetDefaultView(BookmarkListBox.ItemsSource);
            if (view != null)
            {
                view.SortDescriptions.Clear();
                view.SortDescriptions.Add(new SortDescription("Time", ListSortDirection.Ascending));
                if (view is ICollectionViewLiveShaping liveView)
                {
                    liveView.IsLiveSorting = true;
                    liveView.LiveSortingProperties.Add("Time");
                }
            }

            RecManager.Instance.OnRecordingStopped = () =>
            {
                // UIスレッドで実行する必要があるため Dispatcher を使用
                Dispatcher.Invoke(() =>
                {
                    IsRecordMode = false;
                });
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

            FrameRateText.Text = fps.ToString();
            RecManager.Instance.FrameRate = fps;

            if (update)
            {
                Default.Save();
            }
            // ホットキー登録
            RefreshHotkey();
        }

        private DebugWindow? _debugWindow;

        private void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWin = new SettingsWindow();
            settingsWin.Owner = this;
            if (settingsWin.ShowDialog() == true)
            {
                ApplyCurrentSettings();
                StatusText.Text = "⚙️ 設定を更新しました";
                RefreshCanvas();

                if (_setting.IsVoiceEnabled && !_isSpeechInitalized)
                {
                    InitSpeechRecognition();
                }
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
        
        // ファイル -> 終了
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        // ブックマークをクリア
        private void ClearBookmark_Click(object sender, RoutedEventArgs e)
        {
            BookmarkListBox.Items.Clear();
        }

        private void SaveEvidence_Click(object? sender, RoutedEventArgs? e)
        {
            if (RecManager.Instance.Evidence == null) return;

            try
            {
                RecManager.Instance.SaveEvidenceJson();

                StatusText.Text = $"[保存完了] {DateTime.Now:HH:mm:ss} エビデンスを保存しました。";
                MessageBox.Show("エビデンスの内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // エビデンスを開く
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
                try
                {
                    // 1. JSONファイルを読み込む
                    string jsonString = File.ReadAllText(openFileDialog.FileName);

                    // 2. デシリアライズ（復元）
                    var evidence = JsonSerializer.Deserialize<RecEvidence>(jsonString);

                    if (evidence != null)
                    {
                        RecManager.Instance.Evidence = evidence;
                        RecManager.Instance.JsonPath = openFileDialog.FileName;

                        // 3. JSONと同じフォルダ内にある動画ファイルのフルパスを作成
                        var folderPath = Path.GetDirectoryName(openFileDialog.FileName) ?? "";
                        if (!string.IsNullOrEmpty(folderPath))
                        {
                            RecManager.Instance.CurrentFolder = folderPath;
                            string videoPath = Path.Combine(folderPath, evidence?.VideoFileName ?? "");

                            if (File.Exists(videoPath))
                            {
                                // 4. 再生準備
                                VideoPlayer.Source = new Uri(videoPath);
                                VideoPlayer.Visibility = Visibility.Visible;
                                RecordingOverlay.Visibility = Visibility.Collapsed;

                                PlayerPause(true);

                                // 5. UIに情報を反映
                                StatusText.Text = $"読み込み: {evidence?.RecMode} {evidence?.VideoFileName}";

                                // ビュー（ソート用）の再設定
                                BookmarkListBox.ItemsSource = RecManager.Instance.Bookmarks;
                                var view = CollectionViewSource.GetDefaultView(BookmarkListBox.ItemsSource);
                                if (view != null)
                                {
                                    view.SortDescriptions.Clear();
                                    view.SortDescriptions.Add(new SortDescription("Time", ListSortDirection.Ascending));
                                    if (view is ICollectionViewLiveShaping liveView)
                                    {
                                        liveView.IsLiveSorting = true;
                                        liveView.LiveSortingProperties.Add("Time");
                                    }
                                }
                                RefreshBookmarkCanvas();
                                IsRecordMode = false;

                                // 選択状態の管理
                                if (RecManager.Instance.Bookmarks.Count > 0)
                                {
                                    BookmarkListBox.SelectedIndex = 0;
                                    BookmarkListBox.Focus();
                                }

                                // 最大リトライ回数（例：100msごとに50回 ＝ 5秒）
                                int maxRetries = 30;
                                int retryCount = 0;
                                while (retryCount < maxRetries)
                                {
                                    if (VideoPlayer.NaturalDuration.HasTimeSpan)
                                    {
                                        // 準備OK！印を描画してループを抜ける
                                        RefreshBookmarkCanvas();
                                        return;
                                    }
                                    await Task.Delay(100);
                                    retryCount++;
                                }
                                Debug.WriteLine("マーカーの描画をタイムアウトしました。");
                            }
                            else
                            {
                                MessageBox.Show("動画ファイルが見つかりません。");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"読み込みエラー: {ex.Message}");
                }
            }
        }

        private void OpenExport_Click(object sender, RoutedEventArgs e)
        {
            // エクスポート画面を表示
            var exportWin = new ExportWindow();
            exportWin.Owner = this; // 親ウィンドウをセットして中央に表示
            exportWin.ShowDialog();
        }

        private void ClearMarkRectangle()
        {
            DrawingCanvas.Children.Clear();
        }

        private void DrawingCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // 💡 リサイズ時に矩形をクリアして再描画する
            RefreshCanvas();
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

                    StatusText.Text = "📂 フォルダを開きました";
                }
                else
                {
                    StatusText.Text = "❌ 保存先フォルダが見つかりません";
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ フォルダを開けませんでした";
                RecManager.Instance.TraceLogs.Add($"Explorer Error: {ex.Message}");
            }
        }

        private void AddVisualMask(MarkRect rect)
        {
            double containerW = DrawingCanvas.ActualWidth;
            double containerH = DrawingCanvas.ActualHeight;

            if (containerW == 0 || containerH == 0 || VideoPlayer.NaturalVideoWidth == 0) return;

            double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
            double dispW = VideoPlayer.NaturalVideoWidth * ratio;
            double dispH = VideoPlayer.NaturalVideoHeight * ratio;
            double offsetX = (containerW - dispW) / 2.0;
            double offsetY = (containerH - dispH) / 2.0;

            // キャンバス上の実際のピクセル座標に変換
            double rectLeft = (rect.X * dispW) + offsetX;
            double rectTop = (rect.Y * dispH) + offsetY;
            double rectWidth = rect.Width * dispW;
            double rectHeight = rect.Height * dispH;

            // 1. 全体（外枠）のジオメトリ
            RectangleGeometry fullCanvas = new RectangleGeometry(new Rect(0, 0, DrawingCanvas.ActualWidth, DrawingCanvas.ActualHeight));

            // 2. クロップ範囲（内枠）のジオメトリ
            RectangleGeometry cropArea = new RectangleGeometry(new Rect(rectLeft, rectTop, rectWidth, rectHeight));

            // 3. 2つの図形を「除外(Exclude)」で合成（ドーナツ型にする）
            CombinedGeometry maskGeometry = new CombinedGeometry(GeometryCombineMode.Exclude, fullCanvas, cropArea);

            // 4. Path要素としてキャンバスに追加
            System.Windows.Shapes.Path maskPath = new System.Windows.Shapes.Path
            {
                Data = maskGeometry,
                Fill = new SolidColorBrush(Color.FromArgb(200, 40, 40, 40)), // 半透明のグレー
                IsHitTestVisible = false // マウス操作を邪魔しないようにする
            };

            DrawingCanvas.Children.Add(maskPath);
        }

        private void DrawRectOnCanvas(MarkRect rect, bool isCropMode)
        {
            double containerW = DrawingCanvas.ActualWidth;
            double containerH = DrawingCanvas.ActualHeight;

            if (containerW == 0 || containerH == 0 || VideoPlayer.NaturalVideoWidth == 0) return;

            double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
            double dispW = VideoPlayer.NaturalVideoWidth * ratio;
            double dispH = VideoPlayer.NaturalVideoHeight * ratio;
            double offsetX = (containerW - dispW) / 2.0;
            double offsetY = (containerH - dispH) / 2.0;

            var rectBrush = isCropMode ? _setting.CropBrush: _setting.MainBrush;
            var overBrush = isCropMode ? _setting.CropFillBrush : _setting.OverFillBrush;

            // 共通の座標計算
            double rectLeft = (rect.X * dispW) + offsetX;
            double rectTop = (rect.Y * dispH) + offsetY;
            double rectWidth = rect.Width * dispW;
            double rectHeight = rect.Height * dispH;
            bool isCropLocked = RecManager.Instance.Evidence?.IsCropLocked?? false;
            bool canTouch = !isCropMode || !isCropLocked;

            // --- 1. 中央の移動用エリア（透明な塗りつぶし） ---
            var moveArea = new WpfRectangle
            {
                Width = Math.Max(0, rectWidth),
                Height = Math.Max(0, rectHeight),
                Fill = Brushes.Transparent, // 透明だがマウスには反応する
                Cursor = Cursors.SizeAll,
                Tag = rect,
                Stroke = rectBrush,
                StrokeThickness = isCropMode ? 1 : 0,
                IsHitTestVisible = canTouch,
                Opacity = canTouch ? 0.6 : 1.0,
            };

            moveArea.MouseEnter += (s, e) => {
                moveArea.Fill = overBrush; // 設定されているハイライト色（半透明）
            };
            moveArea.MouseLeave += (s, e) => {
                moveArea.Fill = Brushes.Transparent;
            };
            moveArea.MouseLeftButtonUp += EndDrag;

            Canvas.SetLeft(moveArea, rectLeft);
            Canvas.SetTop(moveArea, rectTop);

            moveArea.MouseLeftButtonDown += (s, e) => {
                // ダブルクリック判定
                if (e.ClickCount == 2)
                {
                    ToggleCropArea(rect); // クロップ範囲の切り替えメソッドを呼ぶ
                    e.Handled = true;
                    return;
                }
                // 通常のドラッグ開始処理
                _draggingRect = moveArea;
                _isResizing = false;
                _currentResizeDir = ResizeDirection.Move;
                _lastMousePosition = e.GetPosition(DrawingCanvas);
                moveArea.CaptureMouse();
                e.Handled = true;
            };
            DrawingCanvas.Children.Add(moveArea);

            // --- 2. 4つの辺を作成するための共通ヘルパー関数 ---
            UIElement CreateEdge(double x1, double y1, double x2, double y2, ResizeDirection dir, Brush brush)
            {
                var container = new Canvas();

                // 1. 【当たり判定用】 透明で太い線
                var hitArea = new Line
                {
                    X1 = x1,
                    Y1 = y1,
                    X2 = x2,
                    Y2 = y2,
                    Stroke = Brushes.Transparent, // 💡 見えない
                    StrokeThickness = 10,         // 💡 当たり判定を 10px くらいに広げる
                    Cursor = (dir == ResizeDirection.Left || dir == ResizeDirection.Right) ? Cursors.SizeWE : Cursors.SizeNS,
                    Tag = rect,
                    IsHitTestVisible = canTouch,
                };
                var line = new Line
                {
                    X1 = x1,
                    Y1 = y1,
                    X2 = x2,
                    Y2 = y2,
                    Stroke = brush,
                    StrokeThickness = 2, // 当たり判定のために少し太めにする
                    Cursor = canTouch ? ((dir == ResizeDirection.Left || dir == ResizeDirection.Right) ? Cursors.SizeWE : Cursors.SizeNS) : Cursors.Arrow,
                    Tag = rect
                };

                // --- イベント処理 ---
                hitArea.MouseEnter += (s, e) => line.Stroke = _setting.OverBrush;
                hitArea.MouseLeave += (s, e) => line.Stroke = brush;

                hitArea.MouseLeftButtonDown += (s, e) => {
                    _draggingRect = hitArea; // ドラッグ対象は当たり判定用にする
                    _isResizing = true;
                    _currentResizeDir = dir;
                    _lastMousePosition = e.GetPosition(DrawingCanvas);
                    hitArea.CaptureMouse();
                    e.Handled = true;
                };

                // 前述の MouseUp 処理も追加
                hitArea.MouseLeftButtonUp += EndDrag;

                container.Children.Add(line); // 下に描画
                container.Children.Add(hitArea);    // 上に重ねて当たり判定を確保

                return container;
            }

            // --- 3. 4辺をキャンバスに追加 ---
            DrawingCanvas.Children.Add(CreateEdge(rectLeft, rectTop, rectLeft + rectWidth, rectTop, ResizeDirection.Top, rectBrush));
            DrawingCanvas.Children.Add(CreateEdge(rectLeft, rectTop + rectHeight, rectLeft + rectWidth, rectTop + rectHeight, ResizeDirection.Bottom, rectBrush));
            DrawingCanvas.Children.Add(CreateEdge(rectLeft, rectTop, rectLeft, rectTop + rectHeight, ResizeDirection.Left, rectBrush));
            DrawingCanvas.Children.Add(CreateEdge(rectLeft + rectWidth, rectTop, rectLeft + rectWidth, rectTop + rectHeight, ResizeDirection.Right, rectBrush));

            // --- 4. 情報バッジ（サイズ表示）を追加 ---
            // クロップモードの時だけ表示、あるいは常に表示などお好みで
            if (isCropMode)
            {
                var infoTextBlock = new TextBlock
                {
                    // 実際のピクセルサイズを表示したい場合は rect.Width * NaturalVideoWidth 等を使用
                    //Text = $"{(int)(rect.Width * VideoPlayer.NaturalVideoWidth)} × {(int)(rect.Height * VideoPlayer.NaturalVideoHeight)}",
                    Foreground = Brushes.White,
                    FontSize = 10,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                    Padding = new Thickness(5, 1, 5, 1),
                    VerticalAlignment = VerticalAlignment.Center
                };
                string lockIcon = isCropLocked ? " 🔒" : " 🔓";
                infoTextBlock.Text = $"{(int)(rect.Width * VideoPlayer.NaturalVideoWidth)} × {(int)(rect.Height * VideoPlayer.NaturalVideoHeight)}{lockIcon}";

                var infoBadge = new Border
                {
                    Uid = "RectInfoBadge", // リサイズ中に中身を書き換えるための目印
                    Background = new SolidColorBrush(Color.FromArgb(180, 40, 40, 40)),
                    BorderBrush = rectBrush, // 枠線の色と合わせると統一感が出ます
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(2),
                    Child = infoTextBlock,
                    IsHitTestVisible = true,
                    Margin = new Thickness(5)  // 矩形の角から少し離す
                };

                // --- ホバー時の色の定義 ---
                var normalBackground = new SolidColorBrush(Color.FromArgb(180, 40, 40, 40)); // 通常時
                var hoverBackground = _setting.CropFillBrush;
                infoBadge.ToolTip = isCropLocked ? "クリックして解除（移動・リサイズ可能）" : "クリックして固定（注釈編集を優先）";

                // 初期状態の設定
                infoBadge.Background = normalBackground;

                // --- マウスオーバー（ハイライト）処理 ---
                infoBadge.MouseEnter += (s, e) =>
                {
                    infoBadge.Background = hoverBackground;
                    infoBadge.BorderThickness = new Thickness(1.5); // 枠線を少し太くして強調
                };

                infoBadge.MouseLeave += (s, e) =>
                {
                    infoBadge.Background = normalBackground;
                    infoBadge.BorderThickness = new Thickness(1);   // 元に戻す
                };

                // クリック（マウスダウン）イベント
                if (RecManager.Instance?.Evidence != null)
                    infoBadge.MouseLeftButtonDown += (s, e) => {
                        RecManager.Instance.Evidence.IsCropLocked = !isCropLocked; // ロック状態を反転

                        // UIを即座に更新するために再描画をかける
                        // 描画メソッドを現在の状態で呼び出し直す
                        RefreshCanvas();

                        e.Handled = true; // 他の要素にクリックが伝わるのを防ぐ
                    };
                // ロック中はバッジの背景色を変えて「固定済み」を強調してもOK
                if (isCropLocked)
                {
                    infoBadge.Background = new SolidColorBrush(Color.FromArgb(150, 0, 0, 0)); // ロック中は真っ黒に近く
                }

                Canvas.SetLeft(infoBadge, rectLeft);
                Canvas.SetTop(infoBadge, rectTop);
                Canvas.SetZIndex(infoBadge, 100); // 最前面

                DrawingCanvas.Children.Add(infoBadge);
            }
        }

        private void ToggleCropArea(MarkRect target)
        {
            // 1. すでに自分がクロップ範囲なら解除、そうでなければ設定
            bool isNowSelected = !target.IsCropArea;

            // 2. 「クロップ範囲は1つだけ」ルールを適用
            // (全ブックマークの全矩形を走査)
            foreach (var bm in RecManager.Instance.Evidence?.Bookmarks ?? new())
            {
                foreach (var r in bm.MarkRects)
                {
                    r.IsCropArea = false;
                }
            }

            target.IsCropArea = isNowSelected;

            // 💡 画面を更新して色を変えたり、プレビューを再描画したりする
            RefreshCanvas();
        }

        private void RefreshCanvas()
        {
            DrawingCanvas.Children.Clear();

            var marks = RecManager.Instance.Evidence?.Bookmarks ?? [];

            var cropRect = marks.SelectMany(b => b.MarkRects).FirstOrDefault(r => r.IsCropArea);
            if (cropRect != null)
            {
                DrawRectOnCanvas(cropRect, isCropMode: true);
                AddVisualMask(cropRect);
            }

            double containerW = DrawingCanvas.ActualWidth;
            double containerH = DrawingCanvas.ActualHeight;

            if (containerW == 0 || containerH == 0 || VideoPlayer.NaturalVideoWidth == 0) return;

            double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
            double dispW = VideoPlayer.NaturalVideoWidth * ratio;
            double dispH = VideoPlayer.NaturalVideoHeight * ratio;
            double offsetX = (containerW - dispW) / 2.0;
            double offsetY = (containerH - dispH) / 2.0;

            if (BookmarkListBox.SelectedItem is Bookmark selected)
            {
                // --- 1. 矩形 (MarkRects) の描画 ---
                if (selected.MarkRects != null)
                {
                    foreach (var rect in selected.MarkRects)
                    {
                        if (rect.IsCropArea) continue; // クロップ用は既に描画済みならスキップ

                        DrawRectOnCanvas(rect, isCropMode: false);
                    }
                }

                // --- 2. 吹き出し (Balloons) の描画 ---
                foreach (var note in selected.Balloons)
                {
                    // 座標計算
                    WpfPoint start = new WpfPoint(note.TargetPoint.X * dispW + offsetX, note.TargetPoint.Y * dispH + offsetY);
                    WpfPoint end = new WpfPoint(note.TextPoint.X * dispW + offsetX, note.TextPoint.Y * dispH + offsetY);

                    // 💡 描画はこのメソッド1つに任せる
                    DrawBalloonUI(start, end, note);

                    // 💡 アンカー（つまみ）の生成もここで行う
                    var startAnchor = CreateAnchor(note, true);
                    Canvas.SetLeft(startAnchor, start.X - 10);
                    Canvas.SetTop(startAnchor, start.Y - 10);
                    DrawingCanvas.Children.Add(startAnchor);

                    // 終点アンカーは textBorder 自体にドラッグ機能を持たせるなら不要ですが、
                    // ひとまず透明なつまみとして置くなら以下
                    var endAnchor = CreateAnchor(note, false);
                    Canvas.SetLeft(endAnchor, end.X - 10);
                    Canvas.SetTop(endAnchor, end.Y - 10);
                    DrawingCanvas.Children.Add(endAnchor);
                }
            }
        }

        private void DrawBalloonUI(WpfPoint start, WpfPoint end, BalloonNote note)
        {
            // --- 1. Border の生成 ---
            var textBlock = new TextBlock
            {
                Text = note.Text,
                Foreground = _setting.MainTextBrush,
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 200,
                IsHitTestVisible = false
            };

            var textBorder = new Border
            {
                Background = _setting.MainFillBrush,
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10),
                IsHitTestVisible = true,
                Child = textBlock,
                Tag = note,
                Uid = "Text"
            };

            Canvas.SetZIndex(textBorder, 100);

            // 💡 左上を end に合わせる (半分引かない)
            Canvas.SetLeft(textBorder, end.X);
            Canvas.SetTop(textBorder, end.Y);

            // --- 2. ラインの終点計算 ---
            // 💡 左上に合わせる場合、ラインが Border の角に刺さるように見えるため
            // わずかに内側に食い込ませる(隙間を作らない)ために offset を小さくします
            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            double distance = Math.Sqrt(dx * dx + dy * dy);

            double offset = 2; // 角にピッタリ合わせるための微調整
            double adjustedX2 = end.X;
            double adjustedY2 = end.Y;

            if (distance > offset)
            {
                double ratio = (distance - offset) / distance;
                adjustedX2 = start.X + dx * ratio;
                adjustedY2 = start.Y + dy * ratio;
            }

            // --- 3. 要素の描画 ---
            var line = new Line
            {
                X1 = start.X,
                Y1 = start.Y,
                X2 = adjustedX2,
                Y2 = adjustedY2,
                Stroke = _setting.MainFillBrush,
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection { 4, 2 },
                IsHitTestVisible = false
            };
            Canvas.SetZIndex(line, 0);

            var dot = new Ellipse { Width = 6, Height = 6, Fill = _setting.MainBrush, IsHitTestVisible = false };
            Canvas.SetLeft(dot, start.X - 3);
            Canvas.SetTop(dot, start.Y - 3);
            Canvas.SetZIndex(dot, 1);

            // --- 4. イベント登録 ---
            textBorder.MouseLeftButtonDown += (s, e) => {
                if (e.ClickCount == 2)
                {
                    // 💡 再編集時はこのバルーンを隠す
                    HideBalloonUI(note);
                    ShowBalloonInput(note);
                    e.Handled = true;
                    return;
                }
                _draggingRect = textBorder;
                _lastMousePosition = e.GetPosition(DrawingCanvas);
                textBorder.CaptureMouse();
                e.Handled = true;
            };
            // --- 4. イベント登録 ---
            textBorder.MouseEnter += (s, e) =>
            {
                textBorder.Background = _setting.OverBrush;
                line.Stroke = _setting.OverBrush;
                line.StrokeThickness = 1;
                textBorder.Cursor = Cursors.Hand;
            };
            textBorder.MouseLeave += (s, e) =>
            {
                textBorder.Background = _setting.MainFillBrush;
                line.Stroke = _setting.MainBrush;
                line.StrokeThickness = 1;
            };
            textBorder.MouseLeftButtonUp += EndDrag;

            // --- 5. 追加 ---
            DrawingCanvas.Children.Add(line);
            DrawingCanvas.Children.Add(dot);
            DrawingCanvas.Children.Add(textBorder);
        }

        private void HideBalloonUI(BalloonNote note)
        {
            // Canvasから、このノートに関連する要素をすべて探して削除する
            var targets = DrawingCanvas.Children.OfType<FrameworkElement>()
                            .Where(x => x.Tag == note)
                            .ToList();

            foreach (var target in targets)
            {
                DrawingCanvas.Children.Remove(target);
            }
        }
        // アンカー生成ヘルパー
        Ellipse CreateAnchor(BalloonNote data, bool isTarget)
        {
            var el = new Ellipse
            {
                Width = 20,
                Height = 20,
                Fill = Brushes.Transparent, // 💡 通常は透明、デバッグ時は半透明にすると分かりやすい
                Cursor = Cursors.Hand,
                Tag = data // 参照用
            };

            el.MouseLeftButtonDown += (s, e) => {
                _draggingRect = el;
                // 💡 どちらの点を動かしているか判別するために、Tagに情報を付与するか、
                // 名前（Uidなど）に "Target" か "Text" を入れておくと便利です
                el.Uid = isTarget ? "Target" : "Text";

                _lastMousePosition = e.GetPosition(DrawingCanvas);
                el.CaptureMouse();
                e.Handled = true;
            };
            el.MouseLeftButtonUp += EndDrag; // 前に作った共通の解放関数
            return el;
        }

        void EndDrag(object sender, MouseButtonEventArgs e)
        {
            if (_draggingRect != null)
            {
                _draggingRect.ReleaseMouseCapture();
                _draggingRect = null;
                _isResizing = false;
                _currentResizeDir = ResizeDirection.None;
                e.Handled = true;
            }
        }
        private void DrawingCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Right)
            {
                // ⭐ 最優先：入力中（TextBoxがある）なら、キャンセルして即座に抜ける
                if (_activeBalloonInput != null)
                {
                    CancelBalloonInput();
                    e.Handled = true;
                    return; // 💡 これより下の「削除ロジック」には行かせない！
                }

                // --- 既存の削除ロジック（ここからは入力中でない時だけ実行される） ---
                _isDrawing = false;
                // ... (以下、selectedBm.Balloons から RemoveAt するループなど) ...
            }
            // 💡 吹き出し入力中に、TextBox以外を左クリックしたら確定
            if (_activeBalloonInput != null)
            {
                // クリックされたのが TextBox 本人でないことを確認
                if (!Equals(e.OriginalSource, _activeBalloonInput) && e.LeftButton == MouseButtonState.Pressed)
                {
                    // 保存しておいた座標を使って確定！
                    ConfirmBalloon(_startPoint, _endPoint, _activeBalloonInput.Text);

                    // 💡 確定後は TextBox を消すので、ここでの処理は終了
                    e.Handled = true;
                    return;
                }
            }
            // 💡 1. 右クリック：削除処理
            if (e.ChangedButton == MouseButton.Right)
            {
                _isDrawing = false;
                _currentRectangle = null;

                Bookmark? selectedBm = BookmarkListBox.SelectedItem as Bookmark;
                if (selectedBm == null) return;

                var mousePos = e.GetPosition(DrawingCanvas);

                // 表示領域の計算（比率座標からの復元用）
                double containerW = DrawingCanvas.ActualWidth;
                double containerH = DrawingCanvas.ActualHeight;
                double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
                double dispW = VideoPlayer.NaturalVideoWidth * ratio;
                double dispH = VideoPlayer.NaturalVideoHeight * ratio;
                double offsetX = (containerW - dispW) / 2.0;
                double offsetY = (containerH - dispH) / 2.0;

                // --- 1. 💡 吹き出し (Balloons) の当たり判定 ---
                if (selectedBm.Balloons != null)
                {
                    for (int i = selectedBm.Balloons.Count - 1; i >= 0; i--)
                    {
                        var note = selectedBm.Balloons[i];
                        // テキストラベルの位置とサイズを計算
                        // (UpdateCanvasRects で描画している位置と同じ)
                        double textX = (note.TextPoint.X * dispW) + offsetX;
                        double textY = (note.TextPoint.Y * dispH) + offsetY;

                        // テキストのサイズは動的ですが、判定用におおよそのサイズ（例: 幅100x高さ30）
                        // もしくは厳密にやるなら描画された Border の ActualWidth を使いますが、
                        // ここでは簡易的に「ラベルの開始点から一定範囲」で判定します。
                        Rect textRect = new Rect(textX, textY, 100, 30);

                        if (textRect.Contains(mousePos))
                        {
                            selectedBm.Balloons.RemoveAt(i);
                            RefreshCanvas();
                            e.Handled = true;
                            return;
                        }
                    }
                }

                // --- 2. 従来の矩形 (MarkRects) の当たり判定 ---
                for (int i = selectedBm.MarkRects.Count - 1; i >= 0; i--)
                {
                    var r = selectedBm.MarkRects[i];
                    var absRect = new Rect(r.X * dispW + offsetX, r.Y * dispH + offsetY, r.Width * dispW, r.Height * dispH);

                    if (absRect.Contains(mousePos))
                    {
                        selectedBm.MarkRects.RemoveAt(i);
                        RefreshCanvas();
                        e.Handled = true;
                        return;
                    }
                }
                return;
            }

            // 💡 2. 左クリック時の共通準備
            if (e.ChangedButton == MouseButton.Left)
            {
                _isDrawing = true;
                _startPoint = e.GetPosition(DrawingCanvas);

                if (_isPlaying) PlayPauseButton_Click(null, null);

                // 前回のドラッグ線がもし残っていたら掃除 (2回目以降対策)
                if (_dragLine != null)
                {
                    DrawingCanvas.Children.Remove(_dragLine);
                    _dragLine = null;
                }

                // --- モード分岐 ---

                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    // A. バルーンモード：ここでは「始点」を決めるだけでOK
                    // MouseMove側で _dragLine を作成します
                    return;
                }
                else
                {
                    // B. 通常の矩形描画モード
                    _currentRectangle = new WpfRectangle
                    {
                        Stroke = _setting.OverBrush,
                        StrokeThickness = 2,
                        Fill = _setting.OverFillBrush,
                        IsHitTestVisible = false
                    };

                    Canvas.SetLeft(_currentRectangle, _startPoint.X);
                    Canvas.SetTop(_currentRectangle, _startPoint.Y);
                    DrawingCanvas.Children.Add(_currentRectangle);
                }
            }
        }

        // 2. マウスが移動中：矩形のサイズを更新
        private void DrawingCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_draggingRect != null && _draggingRect.Tag is BalloonNote bNote)
            {
                double containerW = DrawingCanvas.ActualWidth;
                double containerH = DrawingCanvas.ActualHeight;
                if (VideoPlayer.NaturalVideoWidth == 0) return;

                double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
                double dispW = VideoPlayer.NaturalVideoWidth * ratio;
                double dispH = VideoPlayer.NaturalVideoHeight * ratio;

                var currentPos = e.GetPosition(DrawingCanvas);

                // 移動量を計算（正規化座標 0.0-1.0）
                double diffX = (currentPos.X - _lastMousePosition.X) / dispW;
                double diffY = (currentPos.Y - _lastMousePosition.Y) / dispH;

                // 💡 どちらの点を動かしているか判定して更新
                if (_draggingRect.Uid == "Target")
                {
                    bNote.TargetPoint = new WpfPoint(bNote.TargetPoint.X + diffX, bNote.TargetPoint.Y + diffY);
                }
                else if (_draggingRect.Uid == "Text")
                {
                    bNote.TextPoint = new WpfPoint(bNote.TextPoint.X + diffX, bNote.TextPoint.Y + diffY);
                }

                _lastMousePosition = currentPos;
                RefreshCanvas(); // 再描画
                return;
            }
            // --- 1. 矩形の移動・リサイズ処理 ---
            if (_draggingRect != null && _draggingRect.Tag is MarkRect data)
            {
                // 💡 ここで dispW, dispH を再計算する
                double containerW = DrawingCanvas.ActualWidth;
                double containerH = DrawingCanvas.ActualHeight;
                if (VideoPlayer.NaturalVideoWidth == 0) return;

                double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
                double dispW = VideoPlayer.NaturalVideoWidth * ratio;
                double dispH = VideoPlayer.NaturalVideoHeight * ratio;

                var currentPos = e.GetPosition(DrawingCanvas);
                // これで下の diffX, diffY のエラーが消えます
                double diffX = (currentPos.X - _lastMousePosition.X) / dispW;
                double diffY = (currentPos.Y - _lastMousePosition.Y) / dispH;

                if (_isResizing)
                {
                    // 💡 選択された辺に応じて計算を切り替え
                    switch (_currentResizeDir)
                    {
                        case ResizeDirection.Right:
                            data.Width = Math.Max(0.01, data.Width + diffX);
                            break;
                        case ResizeDirection.Left:
                            // 左端を動かすときは、位置(X)を動かしつつ幅を逆方向に調整
                            if (data.Width - diffX > 0.01) { data.X += diffX; data.Width -= diffX; }
                            break;
                        case ResizeDirection.Bottom:
                            data.Height = Math.Max(0.01, data.Height + diffY);
                            break;
                        case ResizeDirection.Top:
                            // 上端を動かすときは、位置(Y)を動かしつつ高さを逆方向に調整
                            if (data.Height - diffY > 0.01) { data.Y += diffY; data.Height -= diffY; }
                            break;
                    }
                }
                else
                {
                    // 通常移動
                    data.X += diffX;
                    data.Y += diffY;
                }

                _lastMousePosition = currentPos;
                RefreshCanvas();
                return;
            }

            // --- 2. バルーン入力キャンセル（右クリック） ---
            if (_activeBalloonInput != null && e.RightButton == MouseButtonState.Pressed)
            {
                CancelBalloonInput();
                e.Handled = true;
                return;
            }

            var overBrush = new SolidColorBrush(_setting.OverColor);

            // --- 3. 新規バルーン描画中のガイド線表示 ---
            if (_isDrawing && Keyboard.Modifiers == ModifierKeys.Control)
            {
                var currentPoint = e.GetPosition(DrawingCanvas);
                UpdateDragLine(currentPoint, overBrush);
                return;
            }

            // --- 4. 吹き出しドラッグ中のガイド線表示 ---
            if (Keyboard.Modifiers == ModifierKeys.Control && e.LeftButton == MouseButtonState.Pressed)
            {
                var currentPoint = e.GetPosition(DrawingCanvas);
                UpdateDragLine(currentPoint, overBrush);
                return;
            }

            // --- 5. 新規矩形の描画（サイズ確定中） ---
            if (e.LeftButton == MouseButtonState.Pressed && _currentRectangle != null)
            {
                var pos = e.GetPosition(DrawingCanvas);
                var x = Math.Min(pos.X, _startPoint.X);
                var y = Math.Min(pos.Y, _startPoint.Y);
                var width = Math.Abs(pos.X - _startPoint.X);
                var height = Math.Abs(pos.Y - _startPoint.Y);

                Canvas.SetLeft(_currentRectangle, x);
                Canvas.SetTop(_currentRectangle, y);
                _currentRectangle.Width = width;
                _currentRectangle.Height = height;
                return;
            }

            // --- 6. ホバー判定など（既存の末尾処理） ---
            // ※ ここは既存の判定ロジックを継続
        }

        // 💡 重複していたガイド線更新を共通化するとスッキリします
        private void UpdateDragLine(WpfPoint currentPoint, Brush brush)
        {
            if (_dragLine == null)
            {
                _dragLine = new Line
                {
                    Stroke = brush,
                    StrokeThickness = 1,
                    StrokeDashArray = new DoubleCollection { 4, 2 },
                    IsHitTestVisible = false
                };
                DrawingCanvas.Children.Add(_dragLine);
            }
            _dragLine.X1 = _startPoint.X;
            _dragLine.Y1 = _startPoint.Y;
            _dragLine.X2 = currentPoint.X;
            _dragLine.Y2 = currentPoint.Y;
        }

        private void ShowBalloonInput(BalloonNote? targetNote = null)
        {
            // 設定から色を取得
            var mainBrush = new SolidColorBrush(_setting.MainTextColor);

            double targetX, targetY;

            if (targetNote != null)
            {
                // --- 💡 再編集モード：UpdateCanvasRects の計算ロジックと同期させる ---
                double containerW = DrawingCanvas.ActualWidth;
                double containerH = DrawingCanvas.ActualHeight;

                // 動画が読み込まれていない、またはサイズがない場合のフォールバック
                if (VideoPlayer.NaturalVideoWidth == 0 || containerW == 0) return;

                double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
                double dispW = VideoPlayer.NaturalVideoWidth * ratio;
                double dispH = VideoPlayer.NaturalVideoHeight * ratio;
                double offsetX = (containerW - dispW) / 2.0;
                double offsetY = (containerH - dispH) / 2.0;

                // 保存されている正規化座標(0.0~1.0)を表示座標に変換
                targetX = (targetNote.TextPoint.X * dispW) + offsetX;
                targetY = (targetNote.TextPoint.Y * dispH) + offsetY;
            }
            else
            {
                // 新規作成モード：マウスを離した座標（_endPoint）を使用
                targetX = _endPoint.X;
                targetY = _endPoint.Y;
            }

            _activeBalloonInput = new System.Windows.Controls.TextBox
            {
                Width = 150,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 14,
                BorderBrush = mainBrush,
                BorderThickness = new Thickness(2),
                Background = new SolidColorBrush(Color.FromArgb(230, 255, 255, 255)), // 少し透ける白
                Text = targetNote?.Text ?? ""
            };

            Canvas.SetLeft(_activeBalloonInput, targetX);
            Canvas.SetTop(_activeBalloonInput, targetY);
            DrawingCanvas.Children.Add(_activeBalloonInput);

            _activeBalloonInput.Focus();
            if (targetNote != null && _activeBalloonInput != null) _activeBalloonInput.SelectAll();

            // 確定処理
            void FinalizeInput()
            {
                if (_activeBalloonInput == null) return;
                string newText = _activeBalloonInput.Text;

                if (targetNote != null)
                {
                    targetNote.Text = newText;
                    CancelBalloonInput();
                    RefreshCanvas();
                }
                else
                {
                    ConfirmBalloon(_startPoint, _endPoint, newText);
                }
            }

            if (_activeBalloonInput != null)
            {
                _activeBalloonInput.LostFocus += (s, e) => FinalizeInput();

                _activeBalloonInput.PreviewKeyDown += (s, e) =>
                {
                    if (e.Key == Key.Enter && Keyboard.Modifiers != ModifierKeys.Shift)
                    {
                        FinalizeInput();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Escape)
                    {
                        if (targetNote == null) CleanupDragLine();
                        CancelBalloonInput();
                        if (targetNote != null) RefreshCanvas(); // 編集キャンセル時は再描画して元に戻す
                    }
                };
            }
        }

        private void CancelBalloonInput()
        {
            // 💡 1. 入力中の TextBox を消す
            if (_activeBalloonInput != null)
            {
                DrawingCanvas.Children.Remove(_activeBalloonInput);
                _activeBalloonInput = null;
            }

            // 💡 2. ガイド用の赤い点線を消す
            CleanupDragLine();

            // 💡 3. _isDrawing フラグを折って、変な残像が出ないようにする
            _isDrawing = false;
        }

        private void CleanupDragLine()
        {
            if (_dragLine != null)
            {
                DrawingCanvas.Children.Remove(_dragLine);
                _dragLine = null;
            }
        }

        private void ConfirmBalloon(WpfPoint start, WpfPoint end, string text)
        {
            var selectedBm = BookmarkListBox.SelectedItem as Bookmark;

            if (!string.IsNullOrWhiteSpace(text))
            {
                // 💡 ここで比率座標に変換してBookMarkに保存するロジックを呼ぶ
                // (矩形と同じように ratio と offset を使って保存)
                // ConfirmBalloon 内の例
                double containerW = DrawingCanvas.ActualWidth;
                double containerH = DrawingCanvas.ActualHeight;
                double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
                double dispW = VideoPlayer.NaturalVideoWidth * ratio;
                double dispH = VideoPlayer.NaturalVideoHeight * ratio;
                double offsetX = (containerW - dispW) / 2.0;
                double offsetY = (containerH - dispH) / 2.0;

                var note = new BalloonNote
                {
                    TargetPoint = new WpfPoint((start.X - offsetX) / dispW, (start.Y - offsetY) / dispH),
                    TextPoint = new WpfPoint((end.X - offsetX) / dispW, (end.Y - offsetY) / dispH),
                    Text = text
                };
                selectedBm?.Balloons.Add(note);
            }

            // 2. 💡 後始末
            if (_activeBalloonInput != null)
            {
                DrawingCanvas.Children.Remove(_activeBalloonInput);
                _activeBalloonInput = null; // nullに戻すことで「入力中ではない」と判定させる
            }
            CleanupDragLine(); // ガイド線を消す
            RefreshCanvas(); // 再描画
        }

        private void DrawingCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // --- 追加：既存の矩形をドラッグ・リサイズしていた場合の終了処理 ---
            if (_draggingRect != null)
            {
                _draggingRect.ReleaseMouseCapture();

                //var mainBrush = new SolidColorBrush(_setting.MainColor);
                // ハイライトを戻す（必要に応じて）
                if (_draggingRect is Line l) l.Stroke = _setting.MainBrush;
                if (_draggingRect is WpfRectangle r) r.Fill = Brushes.Transparent;

                _draggingRect = null;
                _isResizing = false;
                _currentResizeDir = ResizeDirection.None;

                // 移動・リサイズの終了時はここで処理を終える
                return;
            }

            // 1. 描画フラグのチェック（バルーン・矩形共通）
            if (!_isDrawing) return;

            _endPoint = e.GetPosition(DrawingCanvas);
            _isDrawing = false; // 描画終了フラグを立てる

            // 2. 動画がロードされていない場合は何もしない
            if (VideoPlayer.NaturalVideoWidth == 0) return;

            // --- 3. 【共通】ブックマークを探す、または自動作成する ---
            var selectedBm = BookmarkListBox.SelectedItem as Bookmark;
            TimeSpan currentTime = VideoPlayer.Position;

            // 現在選択中のブックマークがない、または再生時間とズレている場合に新規作成
            if (selectedBm == null || Math.Abs(selectedBm.Time.TotalSeconds - currentTime.TotalSeconds) > 0.1)
            {
                // A. 既に登録されているブックマークの中に、現在の再生時間と一致するものがあるか探す
                var existingBm = BookmarkListBox.Items.Cast<Bookmark>()
                    .FirstOrDefault(b => Math.Abs(b.Time.TotalSeconds - currentTime.TotalSeconds) <= 0.1);

                if (existingBm != null)
                {
                    // 一致するものが見つかった場合は、それを選択状態にする
                    selectedBm = existingBm;
                    BookmarkListBox.SelectedItem = selectedBm;
                }
                else
                {
                    // B. 一致するものがなければ、新しくブックマークを登録する
                    Bookmark bookmark = new()
                    {
                        Time = currentTime,
                        Icon = "📝",
                        Note = "Add",
                    };

                    // リストに追加して選択状態にする
                    RecManager.Instance.AddBookmark(bookmark);
                    BookmarkListBox.SelectedItem = bookmark;
                    selectedBm = bookmark;
                    RefreshBookmarkCanvas();
                }
            }

            // --- 4. モード分岐：バルーン（Ctrl）か 矩形か ---

            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                // 💡 対策：もし UpdateCanvasRects で線が消えていたら、Canvas に戻す
                if (_dragLine != null && !DrawingCanvas.Children.Contains(_dragLine))
                {
                    DrawingCanvas.Children.Add(_dragLine);
                }

                ShowBalloonInput();
                return;
            }

            // --- 5. 従来の矩形 (MarkRects) の保存処理 ---
            if (_currentRectangle == null) return;

            double containerW = DrawingCanvas.ActualWidth;
            double containerH = DrawingCanvas.ActualHeight;

            // 動画の表示比率と領域を算出
            double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
            double dispW = VideoPlayer.NaturalVideoWidth * ratio;
            double dispH = VideoPlayer.NaturalVideoHeight * ratio;

            double offsetX = (containerW - dispW) / 2.0;
            double offsetY = (containerH - dispH) / 2.0;

            // 比率座標（0.0〜1.0）へ変換して保存
            var relativeRect = new MarkRect(
                (Math.Min(_startPoint.X, _endPoint.X) - offsetX) / dispW,
                (Math.Min(_startPoint.Y, _endPoint.Y) - offsetY) / dispH,
                Math.Abs(_endPoint.X - _startPoint.X) / dispW,
                Math.Abs(_endPoint.Y - _startPoint.Y) / dispH
            );

            selectedBm.MarkRects.Add(relativeRect);
            RefreshCanvas();
        }

        private void DeleteBookmarkButton_Click(object? sender, RoutedEventArgs? e)
        {
            // 1. 選択されている項目があるかチェック
            if (BookmarkListBox.SelectedItems.Count == 0)
            {
                StatusText.Text = "ℹ️ 削除する項目を選択してください";
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
                // 2. 選択された項目を一度別リストにコピーする
                // (列挙中に元のコレクションを変更するとエラーになるため)
                var bookmarks = BookmarkListBox.SelectedItems.Cast<Bookmark>().ToList();

                foreach (var cp in bookmarks)
                {
                    // 3. データソースから削除
                    RecManager.Instance.Evidence?.Bookmarks.Remove(cp);

                    //BookmarkListBox.Items.Remove(cp); 不要
                }

                // 4. 保存とステータス更新
                NoteEditBox.Text = "";
                StatusText.Text = $"🗑️ {bookmarks.Count} 件削除しました";
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
            // 2. 💡 ここでタイマーを起動！
            StartPlaybackTimer();
        }

        private void VideoPlayer_MediaFailed(object sender, ExceptionRoutedEventArgs e)
        {
            // 1. ステータスバーにエラーを表示
            StatusText.Text = "❌ 再生エラー";

            // 2. ログに詳細を記録（以前の画像で見られた TraceLogs や BookmarkListBox を活用）
            string errorMessage = $"再生に失敗しました: {e.ErrorException.Message}";
            RecManager.Instance.TraceLogs.Add(errorMessage);

            // 3. ユーザーへの通知
            MessageBox.Show(errorMessage, "再生エラー", MessageBoxButton.OK, MessageBoxImage.Error);

            // 4. UIの状態をリセット
            VideoPlayer.Visibility = Visibility.Collapsed;
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
                    StatusText.Text = $"{sm.MoniterName}を選択しました ";
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
                    _selectedRegion = new System.Drawing.Rectangle(
                        (int)rect.X,
                        (int)rect.Y,
                        (int)rect.Width,
                        (int)rect.Height
                    );
                    _rectDeviceName = selectionRect.TargetDeviceName;
                    StatusText.Text = $"矩形確定: {selectionRect.TargetDeviceName} x:{rect.X} y:{rect.Y} w:{rect.Width} h:{rect.Height}";
                }
            }
            // ウインドウ選択モード
            if (ModeComboBox.SelectedIndex == 2)
            {
                var sw = new SelectionWindow();
                sw.ShowDialog();

                _targetWindowHandle = sw.SelectedHWnd;
                StatusText.Text = $"録画対象：{sw.SelectedTitle}";
            }

            // MainWindowを再表示してアクティブにする
            this.Show();
            this.Activate();
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            if (!_isDragging && VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                TimelineSlider.Maximum = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                TimelineSlider.Value = VideoPlayer.Position.TotalSeconds;
            }
        }

        // スライダーを手動で動かしたとき
        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isDragging)
            {
                VideoPlayer.Position = TimeSpan.FromSeconds(TimelineSlider.Value);
                PlayerPause(true);
                BookmarkListBox.SelectedItem = null;
            }
            else
            {
                double currentSec = VideoPlayer.Position.TotalSeconds;

                // 前方誤差 -0.1s ～ 後方誤差 0.3s
                var targetBookmark = RecManager.Instance.Evidence.Bookmarks.FirstOrDefault(bm =>
                {
                    double bmSec = bm.Time.TotalSeconds;
                    double diff = currentSec - bmSec;
                    return diff >= -0.1 && diff < 0.3;
                });

                if (targetBookmark != null)
                {
                    // マッチするブックマークがあれば選択
                    if (BookmarkListBox.SelectedItem != targetBookmark)
                    {
                        _isInternalSelectionChange = true;
                        BookmarkListBox.SelectedItem = targetBookmark;
                        BookmarkListBox.ScrollIntoView(targetBookmark);
                        _isInternalSelectionChange = false;

                        // チェックボックスがONなら停止、OFFならそのまま再生
                        if (StopAtBookmarkCheckBox.IsChecked == true)
                        {
                            PlayerPause(true);
                            StatusText.Text = $"停止（ブックマーク）: {targetBookmark.Time}";
                        }
                    }
                }
                else
                {
                    // 2. どのブックマークの位置でもなければ、選択を解除する
                    if (BookmarkListBox.SelectedItem != null)
                    {
                        BookmarkListBox.SelectedItem = null;
                    }
                }
            }

            RefreshCanvas();
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

            if (RecManager.Instance.Evidence?.Bookmarks == null || !RecManager.Instance.Evidence.Bookmarks.Any()) return;

            // 1. しきい値（0.5秒）以内のものを抽出し
            // 2. 現在地との差が一番小さい順に並べ替え
            // 3. その先頭（最も近いもの）を取得する
            var nearbyBookmark = RecManager.Instance.Evidence.Bookmarks
                .Where(bm => Math.Abs(bm.Time.TotalSeconds - currentValue) < 0.05)
                .OrderBy(bm => Math.Abs(bm.Time.TotalSeconds - currentValue))
                .FirstOrDefault();

            if (nearbyBookmark != null)
            {
                if (BookmarkListBox.SelectedItem != nearbyBookmark)
                {
                    BookmarkListBox.SelectedItem = nearbyBookmark;
                    BookmarkListBox.ScrollIntoView(nearbyBookmark);
                }
            }
        }


        private void PlayPauseButton_Click(object? sender, RoutedEventArgs? e)
        {
            ClearMarkRectangle(); // 再生が始まったら描画を消す
            if (_isPlaying)
            {
                PlayerPause(false);
            }
            else
            {
                // 再生処理
                VideoPlayer.Play();
                PlayPauseIcon.Text = "‖"; // 縦棒二本（一時停止マーク）
                PlayPauseText.Text = "一時停止";
                _isPlaying = true;

                StatusText.Text = "▶ 再生中...";
                //StartPlaybackTimer();
            }
        }

        public void PlayerPause(bool withReflash)
        {
            ClearMarkRectangle(); // 再生が始まったら描画を消す
            if (withReflash) VideoPlayer.Play();

            // 一時停止処理
            VideoPlayer.Pause();
            PlayPauseIcon.Text = "▶"; // 三角マーク
            PlayPauseText.Text = "再生";
            _isPlaying = false;

            StatusText.Text = "⏸ 一時停止中";
        }

        private void StartPlaybackTimer()
        {
            DispatcherTimer playbackTimer = new DispatcherTimer();
            playbackTimer.Interval = TimeSpan.FromMilliseconds(200);
            playbackTimer.Tick += (s, e) =>
            {
                if (VideoPlayer.NaturalDuration.HasTimeSpan)
                {
                    string current = VideoPlayer.Position.ToString(@"mm\:ss");
                    string total = VideoPlayer.NaturalDuration.TimeSpan.ToString(@"mm\:ss");
                    TimeText.Text = $"{current} / {total}";
                    TimelineSlider.Value = VideoPlayer.Position.TotalSeconds;
                }
            };
            playbackTimer.Start();
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
            catch { /* エラー処理 */ }
        }

        // ホットキーが押された時の動作
        private void OnBookmark(object? sender, HotkeyEventArgs e)
        {
            if (RecManager.Instance.IsRecording)
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
            if (RecManager.Instance.IsRecording)
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
            // ブックマーク削除
            //BookmarkListBox.Items.Clear();    不要

            // フラグ更新
            _isRecording = true;

            // イベント登録
            RecManager.Instance.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

            // UIの切り替え
            VideoPlayer.Visibility = Visibility.Collapsed;
            RecordingOverlay.Visibility = Visibility.Visible;
            PlaybackControlsArea.Visibility = Visibility.Collapsed; // 再生系をまとめて隠す
            RecordingTimerArea.Visibility = Visibility.Visible;     // 録画タイマーを表示

            VideoPlayer.Stop();
            //_recordingTimer.Start();

            RecordingIcon.Foreground = Brushes.Black;
            RecordingIcon.Text = "■";
            RecordingText.Text = "録画停止";
        }

        // 録画停止時の処理
        private void OnRecordingStopped()
        {
            // フラグ更新
            _isRecording = false;

            // イベント解除
            RecManager.Instance.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

            RefreshBookmarkCanvas();

            // UIの切り替え
            VideoPlayer.Visibility = Visibility.Visible;
            RecordingOverlay.Visibility = Visibility.Collapsed;
            PlaybackControlsArea.Visibility = Visibility.Visible;
            RecordingTimerArea.Visibility = Visibility.Collapsed;

            StatusText.Text = "保存完了";

            //_recordingTimer.Stop();
            _playerTimer.Start();

            VideoPlayer.Source = new Uri(_currentVideoPath);
            PlayerPause(true);


            RecordingIcon.Foreground = Brushes.Red;
            RecordingIcon.Text = "●";
            RecordingText.Text = "録画開始";

            //AddBookmarkButton.Content = "📌 証跡追加";

        }

        private async void RecordingButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRecording)
            {
                try
                {
                    // 1. 保存されている設定を読み込む

                    string modeName = (ModeComboBox.SelectionBoxItem as string) ?? ModeComboBox.Text;
                    _currentVideoPath = RecManager.Instance.PrepareEvidence(Default.SavePath, modeName);
                    switch (ModeComboBox.SelectedIndex)
                    {
                        case 0: // 全画面
                            RecManager.Instance.StartFullscreenRecording(_currentVideoPath, _fullDeviceName);
                            break;

                        case 1: // 矩形選択
                            if (_selectedRegion == null)
                            {
                                StatusText.Text = "エラー：範囲を先に選択してください";
                                return;
                            }
                            RecManager.Instance.StartRectangleRecording(_currentVideoPath, _rectDeviceName, _selectedRegion);
                            break;

                        case 2: // ウィンドウ選択
                            if (_targetWindowHandle == IntPtr.Zero)
                            {
                                StatusText.Text = "エラー：ウィンドウを先に選択してください";
                                return;
                            }
                            RecManager.Instance.StartWindowRecording(_currentVideoPath, _targetWindowHandle);
                            break;
                    }

                    OnRecordingStarted();

                    StatusText.Text = $"● 録画中: {RecManager.Instance.CurrentVideoName}";

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
                // --- 録画停止の処理 ---
                RecManager.Instance.StopRecording();
                StatusText.Text = "動画を処理中...";
                await Task.Delay(1000);

                OnRecordingStopped();

                if (RecManager.Instance.TraceLogs.Count > 0)
                {
                    taskbarInfo.ProgressState = TaskbarItemProgressState.None;
                    taskbarInfo.ProgressValue = 0;

                    //File.WriteAllLines(_currentLogPath, _recorderManager.TraceLogs, Encoding.UTF8);
                }
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
                if (_previewBitmap == null || _previewBitmap.PixelWidth != width || _previewBitmap.PixelHeight != height)
                {
                    // ScreenRecorderLibのBitmapDataは通常 Bgr32 (24bitの場合は Bgr24)
                    // アルファチャネルを含む場合は Pbgra32 など調整が必要な場合があります
                    _previewBitmap = new WriteableBitmap(width, height, 96, 96, PixelFormats.Bgr32, null);
                    PreviewImage.Source = _previewBitmap;
                }

                _previewBitmap.Lock();
                try
                {
                    // BitmapData.Scan0 (IntPtr) から WriteableBitmap へコピー
                    // 第3引数の bufferSize は Stride * Height で計算
                    int bufferSize = data.Stride * height;

                    _previewBitmap.WritePixels(
                        new Int32Rect(0, 0, width, height),
                        data.Data,
                        bufferSize,
                        data.Stride);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Preview Update Error: {ex.Message}");
                }
                finally
                {
                    _previewBitmap.Unlock();
                }
            }));
        }

        private void BookmarkListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BookmarkListBox.SelectedItem is Bookmark selected)
            {
                if (!_isInternalSelectionChange)
                {
                    VideoPlayer.Position = selected.Time;
                    PlayerPause(true);
                    StatusText.Text = $"Seek: {selected.Time}";
                }

                RefreshCanvas();
                RefreshBookmarkCanvas();
            }
        }
        private void RefreshBookmarkCanvas()
        {
            if (RecManager.Instance.Evidence == null || !VideoPlayer.NaturalDuration.HasTimeSpan) return;
            BookmarkCanvas.Children.Clear();

            const double ThumbWidth = 17.0;
            double canvasWidth = BookmarkCanvas.ActualWidth;
            if (canvasWidth <= 0) return;

            double effectiveWidth = canvasWidth - ThumbWidth;
            if (effectiveWidth <= 0) return;

            double totalSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            if (totalSeconds <= 0) return;

            foreach (var bm in RecManager.Instance.Bookmarks)
            {
                // 時間の割合 (0.0 ～ 1.0)
                //double ratio = bm.Time.TotalSeconds / totalSeconds;
                double ratio = bm.Time.TotalSeconds / totalSeconds;
                ratio = Math.Clamp(ratio, 0.0, 1.0);

                double xPos = (ratio * effectiveWidth) + (ThumbWidth / 2.0);

                bool isSelected = BookmarkListBox.SelectedItems.Contains(bm);
                double r = isSelected ? 1.2 : 1.1;
                // 三角形（▲）の作成
                Polygon triangle = new()
                {
                    Points = [
                        new Point(0, 0),
                        new Point(-6 * r, 10 * r),
                        new Point(6 * r, 10 * r)
                    ],
                    Fill = isSelected ? Brushes.Orange : Brushes.LightGray,
                    Stroke = isSelected ? Brushes.Red : Brushes.Gray,
                    StrokeThickness = 1,
                    Cursor = Cursors.Hand,
                    Tag = bm
                };
                Canvas.SetZIndex(triangle, isSelected ? 10 : 0);

                // イベント登録
                triangle.MouseLeftButtonDown += Marker_MouseLeftButtonDown;

                // 位置調整（xPosが頂点の真ん中にくるように）
                Canvas.SetLeft(triangle, xPos);
                Canvas.SetTop(triangle, 0);

                BookmarkCanvas.Children.Add(triangle);
            }
        }
        private void Marker_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is Bookmark bm)
            {
                // 1. 動画をその時間にジャンプ
                VideoPlayer.Position = bm.Time;

                // 2. (任意) リストボックス等の該当項目を選択状態にする
                BookmarkListBox.SelectedItem = bm;
                BookmarkListBox.ScrollIntoView(bm);

                // Slider側のクリックイベントが動かないように「処理済み」とする
                e.Handled = true;
            }
        }

        private void BookmarkCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            RefreshBookmarkCanvas();
        }
        private bool _isSpeechInitalized = false;
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

        private void AddBookmarkWhileRecording_Click(object sender, RoutedEventArgs? e)
        {
            SoundManager.Instance.PlayShutter();

            Bookmark newBookmark = new()
            {
                Time = RecManager.Instance.CurrentDuration,
                Icon = "📌",
                Note = "Add",
            };

            if (_previewBitmap is not null)
            {
                var path = RecManager.Instance.SaveBackupFromWriteableBitmap(newBookmark, _previewBitmap);
                newBookmark.ImagePath = path;
                StatusText.Text = $"記録 {newBookmark.Time} {newBookmark.Note} SS作成 {path}";
            }
            else
            {
                StatusText.Text = $"記録 {newBookmark.Time} {newBookmark.Note}";
            }
            RecManager.Instance.Bookmarks.Add(newBookmark);
            RefreshBookmarkCanvas();
        }

        private void AddVoiceMemoRecording_Click(object sender, RoutedEventArgs? e)
        {
            Bookmark newBookmark = new()
            {
                Time = RecManager.Instance.CurrentDuration,
                Icon = "🎙",
                Note = "音声入力中",
            };

            RecManager.Instance.Bookmarks.Add(newBookmark);
            RefreshBookmarkCanvas();

            Task.Run(async () => await StartSpeach(newBookmark));

            if (_previewBitmap is not null)
            {
                var path = RecManager.Instance.SaveBackupFromWriteableBitmap(newBookmark, _previewBitmap);
                newBookmark.ImagePath = path;
                StatusText.Text = $"記録 {newBookmark.Time} {newBookmark.Note} SS作成 {path}";
            }
            else
            {
                StatusText.Text = $"記録 {newBookmark.Time} {newBookmark.Note}";
            }
        }

        private void AddBookmarkWhilePlaying_Click(object sender, RoutedEventArgs? e)
        {
            // 選択されているブックマークがある場合は何もしない（誤操作防止）
            if (BookmarkListBox.SelectedItem is Bookmark selected) return;

            Bookmark newBookmark = new()
            {
                Time = VideoPlayer.Position,
                Note = "Add",
                Icon = "📋"
            };

            RecManager.Instance.Bookmarks.Add(newBookmark);
            BookmarkListBox.SelectedItem = newBookmark;
            RefreshBookmarkCanvas();
        }

        private void AddVoiceMemoWhilePlaying_Click(object sender, RoutedEventArgs? e)
        {
            // 選択されているブックマークがある場合は何もしない（誤操作防止）
            if (BookmarkListBox.SelectedItem is Bookmark selected) return;

            Bookmark newBookmark = new()
            {
                Time = VideoPlayer.Position,
                Note = "音声入力中",
                Icon = "🎤"
            };

            RecManager.Instance.Bookmarks.Add(newBookmark);
            BookmarkListBox.SelectedItem = newBookmark;
            RefreshBookmarkCanvas();

            Task.Run(async () => await StartSpeach(newBookmark));
        }

        private async Task StartSpeach(Bookmark bookmark)
        {
            try
            {
                SoundManager.Instance.VoiceStartShutter();

                bookmark.IsListening = true;

                var result = await _winrtRecognizer?.RecognizeAsync();

                if (result.Status == SpeechRecognitionResultStatus.Success)
                {
                    bookmark.Note = result.Text;
                }
                else
                {
                    bookmark.Note = "(認識失敗)";
                }
            }
            catch (Exception)
            {
                bookmark.Note = "(エラー発生)";
            }
            finally
            {
                bookmark.IsListening = false;
                System.Media.SystemSounds.Asterisk.Play();
            }
        }

        private void SwitchToRecordMode_Click(object sender, RoutedEventArgs e)
        {
            IsRecordMode = true;
        }

        private void SwitchToPlayerMode_Click(object sender, RoutedEventArgs e)
        {
            IsRecordMode = false;
        }
    }
}