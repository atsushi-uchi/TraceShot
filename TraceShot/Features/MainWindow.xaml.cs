using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using NHotkey;
using NHotkey.Wpf;
using ScreenRecorderLib;
using System.Diagnostics;
using System.IO;
using System.Media;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Shell;
using System.Windows.Threading;
using TraceShot.Services;
using Brushes = System.Windows.Media.Brushes;
using Drawing = System.Drawing;
using MessageBox = System.Windows.MessageBox;
using WinForms = System.Windows.Forms;
using WpfPoint = System.Windows.Point; // WPFの座標
using WpfRectangle = System.Windows.Shapes.Rectangle;

namespace TraceShot
{
    // 録画範囲の種類を定義
    public enum RecordMode { FullScreen, Region, Window }


    public partial class MainWindow : Window
    {
        private bool _isPlaying = false;
        private bool _isRecording = false;
        private RecorderManager _recorderManager = new RecorderManager();

        private string _currentVideoPath = "";
        private DispatcherTimer _recordingTimer;
        private DispatcherTimer _playerTimer;
        private bool _isDragging = false; // スライダー操作中かどうかの判定

        private Drawing.Rectangle? _selectedRegion = null; // 選択された範囲を保持

        private IntPtr _targetWindowHandle;
        private WpfPoint _startPoint;
        private WpfRectangle _currentRectangle = new ();

        public MainWindow()
        {
            InitializeComponent();
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
        }

        private DebugWindow? _debugWindow;

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

        // 表示 -> 履歴ログをクリア
        private void ClearLog_Click(object sender, RoutedEventArgs e)
        {
            LogListBox.Items.Clear();
        }

        private void SaveEvidence_Click(object sender, RoutedEventArgs e)
        {
            if (_recorderManager.Evidence == null) return;

            try
            {
                _recorderManager.UpdateJson();

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
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
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
                        _recorderManager.Evidence = evidence;
                        _recorderManager.JsonPath = openFileDialog.FileName;
                        ExportPdfButton.IsEnabled = true;
                        ExportExcelButton.IsEnabled = true;

                        // 3. JSONと同じフォルダ内にある動画ファイルのフルパスを作成
                        var folderPath = System.IO.Path.GetDirectoryName(openFileDialog.FileName) ?? "";
                        if (!string.IsNullOrEmpty(folderPath))
                        {
                            _recorderManager.CurrentFolder = folderPath;
                            string videoPath = System.IO.Path.Combine(folderPath, evidence.VideoFileName);

                            if (File.Exists(videoPath))
                            {
                                // 4. 再生準備
                                VideoPlayer.Source = new Uri(videoPath);
                                VideoPlayer.Visibility = Visibility.Visible;
                                RecordingOverlay.Visibility = Visibility.Collapsed;

                                PlayerPause(true);

                                // 5. UIに情報を反映
                                StatusText.Text = $"読み込み: {evidence.WindowTitle} ({evidence.Mode})";

                                // リストボックスにブックマーク一覧を表示（オプション）
                                LogListBox.Items.Clear();
                                if (evidence.Bookmarks != null)
                                {
                                    foreach (var bm in evidence.Bookmarks)
                                    {
                                        LogListBox.Items.Add(bm);
                                        _recorderManager.AddBookmark(bm);
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

        private void ClearMarkRectangle()
        {
            DrawingCanvas.Children.Clear();
        }

        private void DrawingCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // 💡 リサイズ時に矩形をクリアして再描画する
            UpdateCanvasRects();
        }

        private void UpdateCanvasRects()
        {
            DrawingCanvas.Children.Clear();

            // 現在選択されているブックマークがある場合のみ実行
            if (LogListBox.SelectedItem is Bookmark selected && selected.MarkRects != null)
            {
                double containerW = DrawingCanvas.ActualWidth;
                double containerH = DrawingCanvas.ActualHeight;

                if (containerW == 0 || containerH == 0 || VideoPlayer.NaturalVideoWidth == 0) return;

                // 現在のサイズにおける動画の表示領域を算出
                double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
                double dispW = VideoPlayer.NaturalVideoWidth * ratio;
                double dispH = VideoPlayer.NaturalVideoHeight * ratio;
                double offsetX = (containerW - dispW) / 2.0;
                double offsetY = (containerH - dispH) / 2.0;

                foreach (var rect in selected.MarkRects)
                {
                    System.Windows.Shapes.Rectangle visualRect = new System.Windows.Shapes.Rectangle
                    {
                        Stroke = System.Windows.Media.Brushes.Red,
                        StrokeThickness = 2,
                        Width = rect.Width * dispW,
                        Height = rect.Height * dispH,
                        Fill = System.Windows.Media.Brushes.Transparent, // 💡 中を透明にしてクリックしやすくする
                        Tag = rect // 💡 重要：元の相対座標データをTagに紐付けておく
                    };

                    // --- 右クリックメニューの設定 ---
                    ContextMenu menu = new ContextMenu();
                    MenuItem deleteItem = new MenuItem { Header = "この矩形を削除" };

                    // 削除イベントの中身
                    deleteItem.Click += (s, e) => {
                        // 保存リストからこのデータを削除
                        selected.MarkRects.Remove((Rect)visualRect.Tag);
                        // 画面を再描画
                        UpdateCanvasRects();
                    };

                    menu.Items.Add(deleteItem);
                    visualRect.ContextMenu = menu;
                    // ------------------------------

                    Canvas.SetLeft(visualRect, (rect.X * dispW) + offsetX);
                    Canvas.SetTop(visualRect, (rect.Y * dispH) + offsetY);

                    DrawingCanvas.Children.Add(visualRect);
                }
            }
        }

        // 1. マウスが押された：描画開始（始点を記録）
        private void DrawingCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // 一時停止してから描画するのがおすすめ
            if (_isPlaying) PlayPauseButton_Click(null, null);

            _startPoint = e.GetPosition(DrawingCanvas);

            // 新しい矩形を作成
            _currentRectangle = new WpfRectangle
            {
                Stroke = System.Windows.Media.Brushes.Red,
                StrokeThickness = 2,
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(50, 255, 0, 0)) // 塗りつぶし（半透明）
            };

            Canvas.SetLeft(_currentRectangle, _startPoint.X);
            Canvas.SetTop(_currentRectangle, _startPoint.Y);
            DrawingCanvas.Children.Add(_currentRectangle);
        }

        // 2. マウスが移動中：矩形のサイズを更新
        private void DrawingCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _currentRectangle != null)
            {
                var pos = e.GetPosition(DrawingCanvas);

                // 始点と現在のマウス位置から、矩形の「左上」の座標と「幅・高さ」を計算
                var x = Math.Min(pos.X, _startPoint.X);
                var y = Math.Min(pos.Y, _startPoint.Y);
                var width = Math.Abs(pos.X - _startPoint.X);
                var height = Math.Abs(pos.Y - _startPoint.Y);

                Canvas.SetLeft(_currentRectangle, x);
                Canvas.SetTop(_currentRectangle, y);
                _currentRectangle.Width = width;
                _currentRectangle.Height = height;
            }
        }

        private void DrawingCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Bookmark? selectedBm = LogListBox.SelectedItem as Bookmark;

            WpfPoint endPoint = e.GetPosition(DrawingCanvas);

            // 💡 基準をすべて DrawingCanvas のサイズに統一する
            double containerW = DrawingCanvas.ActualWidth;
            double containerH = DrawingCanvas.ActualHeight;

            // 1. 動画がロードされていない場合は何もしない
            if (VideoPlayer.NaturalVideoWidth == 0) return;

            // 2. 現在の再生時間のブックマークを探す、または作成する
            // ※ 現在選択中のブックマークが今の再生時間と一致しない、または未選択の場合
            TimeSpan currentTime = VideoPlayer.Position;

            if (selectedBm == null || Math.Abs(selectedBm.Seconds - currentTime.TotalSeconds) > 0.1)
            {
                // 💡 ここで「証跡追加」ボタンと同じ新規作成ロジックを走らせる
                Bookmark bookmark = new Bookmark
                {
                    Time = currentTime.ToString(@"mm\:ss\.fff"),
                    Seconds = currentTime.TotalSeconds,
                    Note = " - Marking"
                };

                // リストに追加して選択状態にする
                var sorted =_recorderManager.AddBookmark(bookmark);
                LogListBox.Items.Clear();
                foreach (var b in sorted) LogListBox.Items.Add(b);
                LogListBox.SelectedItem = bookmark;
                selectedBm = bookmark;
            }

            // 動画の表示比率と領域を算出
            double ratio = Math.Min(containerW / VideoPlayer.NaturalVideoWidth, containerH / VideoPlayer.NaturalVideoHeight);
            double dispW = VideoPlayer.NaturalVideoWidth * ratio;
            double dispH = VideoPlayer.NaturalVideoHeight * ratio;

            // キャンバス内での動画の左端・上端（余白）を算出
            double offsetX = (containerW - dispW) / 2.0;
            double offsetY = (containerH - dispH) / 2.0;

            // 比率座標（0.0〜1.0）へ変換して保存
            Rect relativeRect = new Rect(
                (Math.Min(_startPoint.X, endPoint.X) - offsetX) / dispW,
                (Math.Min(_startPoint.Y, endPoint.Y) - offsetY) / dispH,
                Math.Abs(endPoint.X - _startPoint.X) / dispW,
                Math.Abs(endPoint.Y - _startPoint.Y) / dispH
            );

            selectedBm.MarkRects.Add(relativeRect);
            UpdateCanvasRects();
        }

        private void DeleteBookmarkButton_Click(object sender, RoutedEventArgs e)
        {
            // 1. 選択されている項目があるかチェック
            var selected = LogListBox.SelectedItem as Bookmark;
            if (selected == null)
            {
                MessageBox.Show("削除する項目を選択してください。", "通知");
                return;
            }

            // 2. 確認ダイアログを表示（誤操作防止）
            var result = MessageBox.Show($"{selected.Time} のブックマークを削除しますか？",
                                         "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                // 3. データソースから削除
                _recorderManager.Evidence.Bookmarks.Remove(selected);

                // 4. UI（ListBox）から削除
                // ItemsSourceを使っている場合は自動で消えますが、Items.Add方式の場合は手動で消します
                LogListBox.Items.Remove(selected);

                // 5. 💡 画像ファイルも削除するかは運用によりますが、
                // 基本的にはディスクに残しておき、手動で整理する方が安全です。

                // 6. 保存とステータス更新
                _recorderManager.SaveEvidenceJson();
                StatusText.Text = $"[削除完了] {selected.Time} の項目を削除しました。";

                // メモ入力欄をクリア
                NoteEditBox.Text = "";
            }
        }
        private void AddBookmarkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isRecording)
            {
                TakeTraceLog(); // 前に準備したログ保存メソッドを呼ぶ
                return;
            }

            if (_recorderManager is null || _recorderManager.Evidence is null) return;

            // 1. 現在の再生時間を取得（秒単位などで丸めるのがおすすめ）
            var currentTime = VideoPlayer.Position;
            string timeStr = currentTime.ToString(@"mm\:ss\.fff");

            // 2. 💡 すでに同じ時間のブックマークがあるかチェック
            bool isDuplicate = _recorderManager.Evidence.Bookmarks.Any(b => b.Time == timeStr);

            if (isDuplicate)
            {
                // 重複している場合は何もしない（またはメッセージを出す）
                return;
            }

            // 3. スクリーンショットの撮影と保存
            string fileName = $"SS_{DateTime.Now:yyyyMMddHHmmss}.png";
            string imagePath = System.IO.Path.Combine(_recorderManager.CurrentFolder, "ScreenShot", fileName);

            // 4. ブックマークリストに追加
            var bookmark = new Bookmark
            {
                Time = timeStr,
                Seconds = currentTime.TotalSeconds,
                Note = "- Add",
                ImagePath = imagePath
            };

            var sorted = _recorderManager.AddBookmark(bookmark);
            LogListBox.Items.Clear();
            foreach (var b in sorted) LogListBox.Items.Add(b);
            
            LogListBox.ScrollIntoView(bookmark);
            StatusText.Text = $"★ 追加 {bookmark.Time} {bookmark.Note}";
            System.Media.SystemSounds.Asterisk.Play();
        }

        private double GetSelectedScale()
        {
            if (ScaleComboBox.SelectedItem is ComboBoxItem item &&
                double.TryParse(item.Tag.ToString(), out double scale))
            {
                return scale;
            }
            return 1.0; // デフォルト
        }

        private async void ExportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            // 録画直後またはJSON読み込み直後のデータを対象にする
            if (_recorderManager.Evidence == null)
            {
                MessageBox.Show("エクスポートするデータがありません。");
                return;
            }

            try
            {
                // 元の再生位置を記憶
                TimeSpan originalPosition = VideoPlayer.Position;
                bool wasPlaying = VideoPlayer.CanPause && VideoPlayer.Position > TimeSpan.Zero;
                PlayerPause(false);
                try
                {
                    foreach (var bm in _recorderManager.Evidence.Bookmarks)
                    {
                        VideoPlayer.Position = TimeSpan.FromSeconds(bm.Seconds);
                        await Task.Delay(500);

                        // 画像を保存し、そのパスを bm.ImagePath に格納するようマネージャー側を調整
                        var scale = GetSelectedScale();
                        var savedPath = _recorderManager.SaveSingleBookmarkImage(bm, VideoPlayer, scale);
                        bm.ImagePath = savedPath;
                    }
                    _recorderManager.ExportToExcel();
                    MessageBox.Show("Excelを出力しました！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"出力エラー: {ex.Message}");
                }
                finally
                {
                    // 元の位置に戻す
                    VideoPlayer.Position = originalPosition;
                    if (wasPlaying)
                    {
                        PlayerPause(true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エクスポート失敗: {ex.Message}");
            }
        }

        private async void ExportPdfButton_Click(object sender, RoutedEventArgs e)
        {
            // 録画直後またはJSON読み込み直後のデータを対象にする
            if (_recorderManager.Evidence == null)
            {
                MessageBox.Show("エクスポートするデータがありません。");
                return;
            }

            try
            {
                // 元の再生位置を記憶
                TimeSpan originalPosition = VideoPlayer.Position;
                bool wasPlaying = VideoPlayer.CanPause && VideoPlayer.Position > TimeSpan.Zero;
                PlayerPause(false);

                try
                {
                    foreach (var bm in _recorderManager.Evidence.Bookmarks)
                    {
                        VideoPlayer.Position = TimeSpan.FromSeconds(bm.Seconds);
                        await Task.Delay(500);

                        // 画像を保存し、そのパスを bm.ImagePath に格納するようマネージャー側を調整
                        var scale = GetSelectedScale();
                        var savedPath = _recorderManager.SaveSingleBookmarkImage(bm, VideoPlayer, scale);
                        bm.ImagePath = savedPath;
                    }
                    var htmlFilePath = _recorderManager.ExportToHtml();
                    var pdfFilePath = System.IO.Path.ChangeExtension(htmlFilePath, ".pdf");
                    await _recorderManager.ExportToPdfAsync(htmlFilePath, pdfFilePath);
                    MessageBox.Show("PDFを出力しました！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"出力エラー: {ex.Message}");
                }
                finally
                {
                    // 元の位置に戻す
                    VideoPlayer.Position = originalPosition;
                    if (wasPlaying)
                    {
                        PlayerPause(true);
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エクスポート失敗: {ex.Message}");
            }
        }

        private void RecordingTimer_Tick(object? sender, EventArgs e)
        {
            // ストップウォッチの経過時間を表示
            BigRecordingTimerText.Text = _recorderManager.RecordingTime;
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
        string _selectedDeviceName = string.Empty;
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
                    _selectedDeviceName = sm.SelectedDeviceName;
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
                    //_targetWindowHandle = IntPtr.Zero;
                    _selectedDeviceName = selectionRect.TargetDeviceName;
                    StatusText.Text = $"矩形確定: {selectionRect.TargetDeviceName} x:{rect.X} y:{rect.Y} w:{rect.Width} h:{rect.Height}";
                }
            }
            // ウインドウ選択モード
            if (ModeComboBox.SelectedIndex == 2)
            {
                // 2. SelectionWindowを作成して表示
                var sw = new SelectionWindow();
                sw.ShowDialog();

                // 3. SelectionWindowが閉じられた後、プロパティから値を取得
                _targetWindowHandle = sw.SelectedHWnd;
                StatusText.Text = $"ウィンドウを選択しました {sw.SelectedWindowTitle}";
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
            }
        }

        // スライダーを掴んだらタイマーを止める（操作しやすくするため）
        private void Slider_DragStarted(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e) => _isDragging = true;

        // 離したら動画の再生位置を確定してタイマー再開
        private void Slider_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            _isDragging = false;
            VideoPlayer.Position = TimeSpan.FromSeconds(TimelineSlider.Value);
        }

        private void PlayPauseButton_Click(object sender, RoutedEventArgs e)
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

        private void PlayerPause(bool withReflash)
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
        private void RegisterHotkey()
        {
            try
            {
                var hotkey = HotkeyRegister.RegisterBookmark(OnBookmark);
                StatusText.Text = $"証跡追加のホットキーを割り当てました {hotkey}";
                AddBookmarkButton.Content += "\n" + hotkey;
            }
            catch (HotkeyAlreadyRegisteredException)
            {
                // 既に登録されている場合の処理
                MessageBox.Show("RegisterHotkey Error");
            }
        }

        // ホットキーが押された時の動作
        private void OnBookmark(object? sender, HotkeyEventArgs e)
        {
            if (_isRecording)
            {
                TakeTraceLog(); // 前に準備したログ保存メソッドを呼ぶ
                e.Handled = true; // 他のアプリにこのキー入力を流さない場合
            }
        }

        private void TakeTraceLog()
        {
            try
            {
                string soundPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\shutter.wav");

                if (File.Exists(soundPath))
                {
                    using (SoundPlayer player = new SoundPlayer(soundPath))
                    {
                        player.Play();
                    }
                }
            }
            catch
            {
                SystemSounds.Beep.Play();
            }

            var bookmark = _recorderManager.AddBookmark();
            if (bookmark is not null)
            {
                LogListBox.Items.Add(bookmark);
                LogListBox.ScrollIntoView(bookmark);
                StatusText.Text = $"★ 記録 {bookmark.Time} {bookmark.Note}";
            }
        }


        // 録画開始時の処理
        private void OnRecordingStarted()
        {
            // ブックマーク削除
            LogListBox.Items.Clear();

            // フラグ更新
            _isRecording = true;

            // イベント登録
            _recorderManager.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

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
            _recorderManager.OnPreviewFrameReceived += RecorderManager_OnPreviewFrameReceived;

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

            ExportPdfButton.IsEnabled = true;
            ExportExcelButton.IsEnabled = true;

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
                    string modeName = (ModeComboBox.SelectionBoxItem as string) ?? ModeComboBox.Text;
                    switch (ModeComboBox.SelectedIndex)
                    {
                        case 0: // 全画面
                            _currentVideoPath = _recorderManager.PrepareEvidence(modeName, string.Empty);
                            _recorderManager.StartFullscreenRecording(_currentVideoPath, _selectedDeviceName);
                            break;

                        case 1: // 矩形選択
                            if (_selectedRegion == null)
                            {
                                StatusText.Text = "エラー：範囲を先に選択してください";
                                return;
                            }
                            _currentVideoPath = _recorderManager.PrepareEvidence(modeName, string.Empty);
                            _recorderManager.StartRectangleRecording(_currentVideoPath, _selectedDeviceName, _selectedRegion);
                            break;

                        case 2: // ウィンドウ選択
                            if (_targetWindowHandle == IntPtr.Zero)
                            {
                                StatusText.Text = "エラー：ウィンドウを先に選択してください";
                                return;
                            }
                            _currentVideoPath = _recorderManager.PrepareEvidence(modeName, "ここにタイトルをいれる");
                            _recorderManager.StartWindowRecording(_currentVideoPath, _targetWindowHandle);
                            break;
                    }

                    // 録画開始ボタンの処理内
                    _recorderManager.OnActualRecordingStarted += (s, e) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            _recordingTimer.Start();
                        });
                    };

                    OnRecordingStarted();

                    RegisterHotkey();
                    StatusText.Text = $"● 録画中: {_recorderManager.CurrentVideoName}";

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
                _recorderManager.StopRecording();
                StatusText.Text = "動画を処理中...";
                await Task.Delay(1000);

                OnRecordingStopped();

                if (_recorderManager.TraceLogs.Count > 0)
                {
                    taskbarInfo.ProgressState = TaskbarItemProgressState.None;
                    taskbarInfo.ProgressValue = 0;

                    //File.WriteAllLines(_currentLogPath, _recorderManager.TraceLogs, Encoding.UTF8);
                }
            }
        }
        private WriteableBitmap? _previewBitmap;
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

        private void LogListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (LogListBox.SelectedItem is Bookmark selected)
                {
                    VideoPlayer.Position = TimeSpan.FromSeconds(selected.Seconds);
                    PlayerPause(true);

                    StatusText.Text = $"ジャンプ: {selected.Time}";

                    // イベント連鎖を防ぐため一度ハンドラを外すか、フラグを使うと安全です
                    NoteEditBox.Text = selected.Note;

                    UpdateCanvasRects();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("エラー: " + ex.Message);
            }
        }

        private void NoteEditBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (LogListBox.SelectedItem is Bookmark selected)
            {
                selected.Note = NoteEditBox.Text;

                // 💡 画面上のリスト表示をリアルタイムに更新（Refresh）
                LogListBox.Items.Refresh();
            }
        }

        private void VideoPlayer_MediaFailed(object sender, ExceptionRoutedEventArgs e)
        {
            MessageBox.Show($"動画の読み込みに失敗しました:\n{e.ErrorException.Message}",
                    "再生エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}