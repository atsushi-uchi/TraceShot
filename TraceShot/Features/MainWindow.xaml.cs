using NHotkey;
using ScreenRecorderLib;
using System.Diagnostics;
using System.IO;
using System.Media;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Shell;
using System.Windows.Threading;
using TraceShot.Models;
using TraceShot.Services;
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
using TextBox = System.Windows.Controls.TextBox;
using WpfPoint = System.Windows.Point; // WPFの座標
using WpfRectangle = System.Windows.Shapes.Rectangle;

namespace TraceShot.Features
{
    // 録画範囲の種類を定義
    public enum RecordMode { FullScreen, Region, Window }


    public partial class MainWindow : Window
    {
        private SettingsService _setting = SettingsService.Instance;
        private bool _isPlaying = false;
        private bool _isRecording = false;
        public RecorderManager RecorderMgr { get; private set; }  = new ();

        private string _currentVideoPath = "";
        private DispatcherTimer _recordingTimer;
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
            _recordingTimer = new DispatcherTimer();
            _recordingTimer.Interval = TimeSpan.FromMilliseconds(500); // 0.5秒ごとに更新
            _recordingTimer.Tick += RecordingTimer_Tick;

            // 録画開始ボタンの処理内
            RecorderMgr.OnActualRecordingStarted += (s, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    _recordingTimer.Start();
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
            RecorderMgr.FrameRate = fps;

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
            if (RecorderMgr.Evidence == null) return;

            try
            {
                RecorderMgr.UpdateJson();

                StatusText.Text = $"[保存完了] {DateTime.Now:HH:mm:ss} エビデンスを保存しました。";
                MessageBox.Show("エビデンスの内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // エビデンスを開く
        private void OpenEvidence_Click(object sender, RoutedEventArgs e)
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
                    var evidence = JsonSerializer.Deserialize<RecordingEvidence>(jsonString);

                    if (evidence != null)
                    {
                        RecorderMgr.Evidence = evidence;
                        RecorderMgr.SyncBookmark();
                        RecorderMgr.JsonPath = openFileDialog.FileName;

                        // 3. JSONと同じフォルダ内にある動画ファイルのフルパスを作成
                        var folderPath = Path.GetDirectoryName(openFileDialog.FileName) ?? "";
                        if (!string.IsNullOrEmpty(folderPath))
                        {
                            RecorderMgr.CurrentFolder = folderPath;
                            string videoPath = Path.Combine(folderPath, evidence?.VideoFileName ?? "");

                            if (File.Exists(videoPath))
                            {
                                // 4. 再生準備
                                VideoPlayer.Source = new Uri(videoPath);
                                VideoPlayer.Visibility = Visibility.Visible;
                                RecordingOverlay.Visibility = Visibility.Collapsed;

                                PlayerPause(true);

                                // 5. UIに情報を反映
                                StatusText.Text = $"読み込み: {evidence?.Mode} {evidence?.VideoFileName}";

                                // リストボックスにブックマーク一覧を表示
                                BookmarkListBox.Items.Clear();
                                if (evidence?.Bookmarks != null)
                                {
                                    foreach (var bm in evidence.Bookmarks)
                                    {
                                        BookmarkListBox.Items.Add(bm);
                                    }
                                    // --- 追加：最初のアイテムを選択する ---
                                    if (BookmarkListBox.Items.Count > 0)
                                    {
                                        BookmarkListBox.SelectedIndex = 0; // 0番目を選択
                                        BookmarkListBox.Focus();           // 必要に応じてフォーカスを当てる
                                    }
                                }
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
                RecorderMgr.TraceLogs.Add($"Explorer Error: {ex.Message}");
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

            // --- 1. 中央の移動用エリア（透明な塗りつぶし） ---
            var moveArea = new WpfRectangle
            {
                Width = Math.Max(0, rectWidth),
                Height = Math.Max(0, rectHeight),
                Fill = Brushes.Transparent, // 透明だがマウスには反応する
                Cursor = Cursors.SizeAll,
                Tag = rect, // MarkRectを保持
                // クロップモードなら少し枠線を太く・目立たせる
                Stroke = rectBrush,
                StrokeThickness = isCropMode ? 1 : 0
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
                    Tag = rect
                };
                var line = new Line
                {
                    X1 = x1,
                    Y1 = y1,
                    X2 = x2,
                    Y2 = y2,
                    Stroke = brush,
                    StrokeThickness = 2, // 当たり判定のために少し太めにする
                    Cursor = (dir == ResizeDirection.Left || dir == ResizeDirection.Right) ? Cursors.SizeWE : Cursors.SizeNS,
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
            // 上辺
            DrawingCanvas.Children.Add(CreateEdge(rectLeft, rectTop, rectLeft + rectWidth, rectTop, ResizeDirection.Top, rectBrush));
            // 下辺
            DrawingCanvas.Children.Add(CreateEdge(rectLeft, rectTop + rectHeight, rectLeft + rectWidth, rectTop + rectHeight, ResizeDirection.Bottom, rectBrush));
            // 左辺
            DrawingCanvas.Children.Add(CreateEdge(rectLeft, rectTop, rectLeft, rectTop + rectHeight, ResizeDirection.Left, rectBrush));
            // 右辺
            DrawingCanvas.Children.Add(CreateEdge(rectLeft + rectWidth, rectTop, rectLeft + rectWidth, rectTop + rectHeight, ResizeDirection.Right, rectBrush));
        }

        private void ToggleCropArea(MarkRect target)
        {
            // 1. すでに自分がクロップ範囲なら解除、そうでなければ設定
            bool isNowSelected = !target.IsCropArea;

            // 2. 「クロップ範囲は1つだけ」ルールを適用
            // (全ブックマークの全矩形を走査)
            foreach (var bm in RecorderMgr?.Evidence?.Bookmarks ?? new())
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

            var marks = RecorderMgr?.Evidence?.Bookmarks?? new();

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

            if (BookmarkListBox.SelectedItem is BookMark selected)
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

                BookMark? selectedBm = BookmarkListBox.SelectedItem as BookMark;
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
            var selectedBm = BookmarkListBox.SelectedItem as BookMark;

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
            var selectedBm = BookmarkListBox.SelectedItem as BookMark;
            TimeSpan currentTime = VideoPlayer.Position;

            // 現在選択中のブックマークがない、または再生時間とズレている場合に新規作成
            if (selectedBm == null || Math.Abs(selectedBm.Time.TotalSeconds - currentTime.TotalSeconds) > 0.1)
            {
                // A. 既に登録されているブックマークの中に、現在の再生時間と一致するものがあるか探す
                var existingBm = BookmarkListBox.Items.Cast<BookMark>()
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
                    BookMark bookmark = new()
                    {
                        Time = currentTime,
                        Note = " - Marking"
                    };

                    // リストに追加して選択状態にする
                    var sorted = RecorderMgr.AddBookmark(bookmark);
                    BookmarkListBox.Items.Clear();
                    foreach (var b in sorted) BookmarkListBox.Items.Add(b);
                    BookmarkListBox.SelectedItem = bookmark;
                    selectedBm = bookmark;
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
                var bookmarks = BookmarkListBox.SelectedItems.Cast<BookMark>().ToList();

                foreach (var cp in bookmarks)
                {
                    // 3. データソースから削除
                    RecorderMgr.Evidence?.Bookmarks.Remove(cp);

                    BookmarkListBox.Items.Remove(cp);
                }

                // 4. 保存とステータス更新
                NoteEditBox.Text = "";
                StatusText.Text = $"🗑️ {bookmarks.Count} 件削除しました";
            }
        }

        private void AddBookmarkButton_Click(object sender, RoutedEventArgs e)
        {
            AddBookmark();
        }


        private void RecordingTimer_Tick(object? sender, EventArgs e)
        {
            // ストップウォッチの経過時間を表示
            BigRecordingTimerText.Text = RecorderMgr.RecordingTime;
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
            RecorderMgr.TraceLogs.Add(errorMessage);

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
            RefreshCanvas();
        }

        // スライダーを掴んだらタイマーを止める（操作しやすくするため）
        private void Slider_DragStarted(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e) => _isDragging = true;

        // 離したら動画の再生位置を確定してタイマー再開
        private void Slider_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            _isDragging = false;
            VideoPlayer.Position = TimeSpan.FromSeconds(TimelineSlider.Value);
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
                // 設定値を取得
                Key key = (Key)Default.HotkeyKey;
                ModifierKeys mod = (ModifierKeys)Default.HotkeyMod;

                // 再登録（内部で AddOrReplace が走るので古いものは上書きされる）
                HotkeyRegister.RegisterBookmark(key, mod, OnBookmark);
            }
            catch { /* エラー処理 */ }
        }

        // ホットキーが押された時の動作
        private void OnBookmark(object? sender, HotkeyEventArgs e)
        {
            AddBookmark();
            e.Handled = true; // 他のアプリにこのキー入力を流さない場合
        }

        private void AddBookmark()
        {
            if (_isRecording)
            {
                TakeBookmark(); // 前に準備したログ保存メソッドを呼ぶ
                return;
            }

            if (RecorderMgr is null || RecorderMgr.Evidence is null) return;

            // 1. 現在の再生時間を取得（秒単位などで丸めるのがおすすめ）
            var currentTime = VideoPlayer.Position;

            // 2. 💡 すでに同じ時間のブックマークがあるかチェック
            bool isDuplicate = RecorderMgr.Evidence.Bookmarks.Any(b => b.Time == currentTime);

            if (isDuplicate)
            {
                // 重複している場合は何もしない（またはメッセージを出す）
                return;
            }

            // 3. スクリーンショットの撮影と保存
            string fileName = $"SS_{DateTime.Now:yyyy-MM-dd_HHmmss_fff}.png";
            string imagePath = Path.Combine(RecorderMgr.CurrentFolder, "ScreenShot", fileName);

            // 4. ブックマークリストに追加
            var bookmark = new BookMark
            {
                Time = currentTime,
                Note = "- Add",
                ImagePath = imagePath
            };

            var sorted = RecorderMgr.AddBookmark(bookmark);
            BookmarkListBox.Items.Clear();
            foreach (var b in sorted) BookmarkListBox.Items.Add(b);

            BookmarkListBox.ScrollIntoView(bookmark);
            StatusText.Text = $"★ 追加 {bookmark.Time} {bookmark.Note}";
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void TakeBookmark()
        {
            try
            {
                string soundPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\shutter.wav");

                if (File.Exists(soundPath))
                {
                    using SoundPlayer player = new SoundPlayer(soundPath);
                    player.Play();
                }
            }
            catch
            {
                SystemSounds.Beep.Play();
            }

            var bookmark = RecorderMgr.AddBookmark();
            if (bookmark is not null)
            {
                BookmarkListBox.Items.Add(bookmark);
                BookmarkListBox.ScrollIntoView(bookmark);
                StatusText.Text = $"記録 {bookmark.Time} {bookmark.Note}";

                // 💡 録画中ならプレビュー用ビットマップをそのまま使う
                if (_previewBitmap != null)
                {
                    // 念のため、この時点でのビットマップの状態を保存
                    RecorderMgr.SaveBackupFromWriteableBitmap(bookmark, _previewBitmap);
                }
            }
        }


        // 録画開始時の処理
        private void OnRecordingStarted()
        {
            // ブックマーク削除
            BookmarkListBox.Items.Clear();

            // フラグ更新
            _isRecording = true;

            // イベント登録
            RecorderMgr.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

            // UIの切り替え
            VideoPlayer.Visibility = Visibility.Collapsed;
            RecordingOverlay.Visibility = Visibility.Visible;
            PlaybackControlsArea.Visibility = Visibility.Collapsed; // 再生系をまとめて隠す
            RecordingTimerArea.Visibility = Visibility.Visible;     // 録画タイマーを表示

            VideoPlayer.Stop();
            _recordingTimer.Start();

            StartStopIcon.Foreground = Brushes.Black;
            StartStopIcon.Text = "■";
            StartStopText.Text = "録画停止";
        }

        // 録画停止時の処理
        private void OnRecordingStopped()
        {
            // フラグ更新
            _isRecording = false;

            // イベント解除
            RecorderMgr.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

            // UIの切り替え
            VideoPlayer.Visibility = Visibility.Visible;
            RecordingOverlay.Visibility = Visibility.Collapsed;
            PlaybackControlsArea.Visibility = Visibility.Visible;
            RecordingTimerArea.Visibility = Visibility.Collapsed;

            StatusText.Text = "保存完了";

            _recordingTimer.Stop();
            _playerTimer.Start();

            VideoPlayer.Source = new Uri(_currentVideoPath);
            PlayerPause(true);


            StartStopIcon.Foreground = Brushes.Red;
            StartStopIcon.Text = "●";
            StartStopText.Text = "録画開始";

            AddBookmarkButton.Content = "📌 証跡追加";
        }

        private async void StartStopButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRecording)
            {
                try
                {
                    // 1. 保存されている設定を読み込む

                    string modeName = (ModeComboBox.SelectionBoxItem as string) ?? ModeComboBox.Text;
                    _currentVideoPath = RecorderMgr.PrepareEvidence(Default.SavePath, modeName);
                    switch (ModeComboBox.SelectedIndex)
                    {
                        case 0: // 全画面
                            RecorderMgr.StartFullscreenRecording(_currentVideoPath, _fullDeviceName);
                            break;

                        case 1: // 矩形選択
                            if (_selectedRegion == null)
                            {
                                StatusText.Text = "エラー：範囲を先に選択してください";
                                return;
                            }
                            RecorderMgr.StartRectangleRecording(_currentVideoPath, _rectDeviceName, _selectedRegion);
                            break;

                        case 2: // ウィンドウ選択
                            if (_targetWindowHandle == IntPtr.Zero)
                            {
                                StatusText.Text = "エラー：ウィンドウを先に選択してください";
                                return;
                            }
                            RecorderMgr.StartWindowRecording(_currentVideoPath, _targetWindowHandle);
                            break;
                    }

                    OnRecordingStarted();

                    StatusText.Text = $"● 録画中: {RecorderMgr.CurrentVideoName}";

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
                RecorderMgr.StopRecording();
                StatusText.Text = "動画を処理中...";
                await Task.Delay(1000);

                OnRecordingStopped();

                if (RecorderMgr.TraceLogs.Count > 0)
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
            try
            {
                if (BookmarkListBox.SelectedItem is BookMark selected)
                {
                    VideoPlayer.Position = selected.Time;
                    PlayerPause(true);

                    StatusText.Text = $"ジャンプ: {selected.Time}";

                    // イベント連鎖を防ぐため一度ハンドラを外すか、フラグを使うと安全です
                    NoteEditBox.Text = selected.Note;

                    RefreshCanvas();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("エラー: " + ex.Message);
            }
        }

        private void NoteEditBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (BookmarkListBox.SelectedItem is BookMark selected)
            {
                selected.Note = NoteEditBox.Text;

                // 💡 画面上のリスト表示をリアルタイムに更新（Refresh）
                BookmarkListBox.Items.Refresh();
            }
        }
    }
}