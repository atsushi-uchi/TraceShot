using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TraceShot.Models;
using TraceShot.Services;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Canvas = System.Windows.Controls.Canvas;
using Color = System.Windows.Media.Color;
using Cursors = System.Windows.Input.Cursors;
using Line = System.Windows.Shapes.Line;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Point = System.Windows.Point;
using Rectangle = System.Windows.Shapes.Rectangle;

namespace TraceShot.Controls
{
    public class RegionUI : Canvas
    {
        public event Action<RegionUI>? DoubleClicked;
        public event Action<RegionUI>? RequestDelete;
        public event Action<RegionUI>? RequestOcrScan;
        public event Action<RegionUI>? Changed;

        private enum ResizeDirection { None, Left, Right, Top, Bottom, Move }


        // データモデルへの参照
        public EvidenceRect Data { get; }

        // 内部パーツ
        private readonly Rectangle _fillArea;
        private readonly List<UIElement> _edges = new();
        private static readonly Brush _defaultHatchBrush = CreateHatchBrush();

        private IInputElement? _draggingRect = null;
        private ResizeDirection _currentResizeDir = ResizeDirection.None;
        private Point _lastMousePosition;

        // フラグ
        private bool _isResizing = false;
        //private bool _isCropLocked = false;
        private readonly List<UIElement> _hitTestElements = new();
        private TextBlock _infoTextBlock;
        private Border _infoBadge;

        public RegionUI(EvidenceRect data)
        {
            Data = data;

            // コンテキストメニュー
            InitializeContextMenu();

            // 1. 中央の塗りつぶしエリア（背面）
            _fillArea = new Rectangle
            {
                Fill = Brushes.Transparent,
                Cursor = Cursors.SizeAll,
                IsHitTestVisible = true,
            };

            InitializeFillArea();

            CropAreaSetting();

            ApplyDataToUI();

            MouseLeftButtonUp += EndDrag;
        }

        private void CropAreaSetting()
        {
            _infoTextBlock = new TextBlock
            {
                // 実際のピクセルサイズを表示したい場合は rect.Width * NaturalVideoWidth 等を使用
                //Text = $"{(int)(rect.Width * VideoPlayer.NaturalVideoWidth)} × {(int)(rect.Height * VideoPlayer.NaturalVideoHeight)}",
                Foreground = Brushes.White,
                FontSize = 10,
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                Padding = new Thickness(5, 1, 5, 1),
                VerticalAlignment = VerticalAlignment.Center
            };

            _infoBadge = new Border
            {
                Uid = "RectInfoBadge", // リサイズ中に中身を書き換えるための目印
                //Background = new SolidColorBrush(Color.FromArgb(180, 40, 40, 40)),
                BorderBrush = SettingsService.Instance.CropBrush, // 枠線の色と合わせると統一感が出ます
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(2),
                Child = _infoTextBlock,
                //IsHitTestVisible = true,
                Margin = new Thickness(5)  // 矩形の角から少し離す
            };

            // --- ホバー時の色の定義 ---
            var normalBackground = new SolidColorBrush(Color.FromArgb(180, 40, 40, 40)); // 通常時
            var hoverBackground = SettingsService.Instance.CropFillBrush;

            // 初期状態の設定
            _infoBadge.Background = normalBackground;

            // --- マウスオーバー（ハイライト）処理 ---
            _infoBadge.MouseEnter += (s, e) =>
            {
                _infoBadge.Background = hoverBackground;
                _infoBadge.BorderThickness = new Thickness(1.5); // 枠線を少し太くして強調
            };

            _infoBadge.MouseLeave += (s, e) =>
            {
                _infoBadge.Background = normalBackground;
                _infoBadge.BorderThickness = new Thickness(1);   // 元に戻す
            };

            // クリック（マウスダウン）イベント
            _infoBadge.MouseLeftButtonDown += (s, e) => {
                RecService.Instance.IsCropLocked = !RecService.Instance.IsCropLocked; // ロック状態を反転

                // UIを即座に更新するために再描画をかける
                // 描画メソッドを現在の状態で呼び出し直す
                UpdateCropUI();
                ApplyDataToUI();
                e.Handled = true; // 他の要素にクリックが伝わるのを防ぐ
            };
            // ロック中はバッジの背景色を変えて「固定済み」を強調してもOK
            if (RecService.Instance.IsCropLocked)
            {
                _infoBadge.Background = new SolidColorBrush(Color.FromArgb(150, 0, 0, 0)); // ロック中は真っ黒に近く
            }
            SetZIndex(_infoBadge, 100); // 最前面

            this.Children.Add(_infoBadge);

            UpdateCropUI();
        }

        public void UpdateCropUI()
        {
            if (Data.IsCropArea)
            {
                _infoBadge.Visibility = Visibility.Visible;

                // 2. 背景色の更新
                _infoBadge.Background = RecService.Instance.IsCropLocked
                    ? new SolidColorBrush(Color.FromArgb(150, 0, 0, 0))
                    : new SolidColorBrush(Color.FromArgb(180, 40, 40, 40));

                // ロック中：バッジ以外はマウスを透過させる
                _fillArea.IsHitTestVisible = !RecService.Instance.IsCropLocked;
                foreach (var element in _hitTestElements)
                {
                    element.IsHitTestVisible = !RecService.Instance.IsCropLocked;
                }

                // バッジだけは常にクリック可能にしておく
                _infoBadge.IsHitTestVisible = true;

                // 1. テキストとアイコンの更新
                _infoTextBlock.Text = RecService.Instance.IsCropLocked ? "ロック中 🔒" : "アンロック 🔓";
                _infoBadge.ToolTip = RecService.Instance.IsCropLocked ? "クリックして解除（移動・リサイズ可能）" : "クリックして固定（注釈編集を優先）";
            }
            else
            {
                _infoBadge.Visibility = Visibility.Collapsed;
            }
        }

        private void InitializeContextMenu()
        {
            var menu = new ContextMenu();

            // 1. マスク処理（黒塗）の切り替え
            var maskItem = new MenuItem { 
                Header = Data.IsMasked ? "マスク解除" : "マスク処理（黒塗）",
                Icon = Data.IsMasked ? "🔓" : "㊙️",
            };
            maskItem.Click += (s, e) =>
            {
                if (Data.IsMasked)
                {
                    Data.IsMasked = false;
                    maskItem.Icon = "㊙️";
                    maskItem.Header = "マスク処理（黒塗）";
                    _fillArea.Fill = Brushes.Transparent;
                }
                else
                {
                    Data.IsMasked = true;
                    maskItem.Icon = "🔓";
                    maskItem.Header = "マスク解除";
                    _fillArea.Fill = _defaultHatchBrush;
                }
                ApplyDataToUI(); // 描画を更新（点線やハッチングの切り替え）
            };

            // 2. 削除
            var deleteItem = new MenuItem { Header = "注釈を削除", Icon = "❌" };
            deleteItem.Click += (s, e) =>
            {
                (this.Parent as Canvas)?.Children.Remove(this);
                RequestDelete?.Invoke(this);
            };

            // 3. OCR（必要であれば）
            var ocrItem = new MenuItem { Header = "OCRで読み取る", Icon = "🔍" };
            ocrItem.Click += async (s, e) =>
            {
                // MainWindow 等に定義した OCR 処理を呼び出す
                RequestOcrScan?.Invoke(this);
            };

            menu.Items.Add(deleteItem);
            menu.Items.Add(new Separator());
            menu.Items.Add(maskItem);
            menu.Items.Add(new Separator());
            menu.Items.Add(ocrItem);

            // RegionUI 自体の右クリックメニューとして設定
            this.ContextMenu = menu;
        }

        private void InitializeFillArea()
        {
            _fillArea.Fill = Data.IsMasked ? _defaultHatchBrush : Brushes.Transparent;

            // イベントの紐付け
            _fillArea.MouseLeftButtonDown += OnFillAreaMouseDown;
            _fillArea.MouseEnter += (s, e) => {
                // マスク中でなければ、ハイライト色を表示するなどの演出が可能
                _fillArea.Fill = Data.IsCropArea
                    ? SettingsService.Instance.CropFillBrush
                    : SettingsService.Instance.OverFillBrush;
            };
            _fillArea.MouseLeave += (s, e) => {
                if (_draggingRect == _fillArea) return;
                _fillArea.Fill = Data.IsMasked ? _defaultHatchBrush : Brushes.Transparent;
            };
            this.Children.Add(_fillArea);
        }


        public void ApplyDataToUI()
        {
            var parentCanvas = this.Parent as Canvas;
            if (parentCanvas == null) return;

            // --- ここからビデオ表示領域の計算 ---
            double containerW = parentCanvas.ActualWidth;
            double containerH = parentCanvas.ActualHeight;

            // ビデオの元のサイズ（MainWindowやServiceから取得）
            double videoW = 1920; // 仮：VideoPlayer.NaturalVideoWidth
            double videoH = 1080; // 仮：VideoPlayer.NaturalVideoHeight

            // キャンバス内でのビデオ表示倍率
            double ratio = Math.Min(containerW / videoW, containerH / videoH);

            // 実際にビデオが表示されているサイズ
            double dispW = videoW * ratio;
            double dispH = videoH * ratio;

            // 黒い余白のサイズ
            double offsetX = (containerW - dispW) / 2.0;
            double offsetY = (containerH - dispH) / 2.0;

            // --- モデル(0.0~1.0)を実際のピクセル座標に変換 ---
            double w = Data.Width * dispW;
            double h = Data.Height * dispH;
            double left = (Data.X * dispW) + offsetX;
            double top = (Data.Y * dispH) + offsetY;

            // UI要素に反映
            this.Width = Math.Max(0, w);
            this.Height = Math.Max(0, h);
            Canvas.SetLeft(this, left);
            Canvas.SetTop(this, top);

            // 内部パーツの更新
            _fillArea.Width = this.Width;
            _fillArea.Height = this.Height;

            RefreshEdges(this.Width, this.Height);

            Changed?.Invoke(this);
        }

        private void RefreshEdges(double w, double h)
        {
            // 既存の辺をクリア（_fillArea 以外の辺パーツを削除）
            foreach (var edge in _edges)
            {
                this.Children.Remove(edge);
            }
            _edges.Clear();
            _hitTestElements.Clear();

            // 枠線の色を決定
            var rectBrush = Data.IsCropArea ? SettingsService.Instance.CropBrush : SettingsService.Instance.MainBrush;

            // 4つの辺を生成して追加
            // 上辺
            _edges.Add(CreateEdge(0, 0, w, 0, ResizeDirection.Top, rectBrush));
            // 下辺
            _edges.Add(CreateEdge(0, h, w, h, ResizeDirection.Bottom, rectBrush));
            // 左辺
            _edges.Add(CreateEdge(0, 0, 0, h, ResizeDirection.Left, rectBrush));
            // 右辺
            _edges.Add(CreateEdge(w, 0, w, h, ResizeDirection.Right, rectBrush));

            foreach (var edge in _edges)
            {
                this.Children.Add(edge);
            }
        }

        private UIElement CreateEdge(double x1, double y1, double x2, double y2, ResizeDirection dir, Brush brush)
        {
            var container = new Canvas();
            bool canTouch = !RecService.Instance.IsCropLocked || !Data.IsCropArea;
            // 1. 【当たり判定用】 透明で太い線
            var hitArea = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = Brushes.Transparent,
                StrokeThickness = 10, // マウスで掴みやすくする
                Cursor = (dir == ResizeDirection.Left || dir == ResizeDirection.Right) ? Cursors.SizeWE : Cursors.SizeNS,
                Tag = Data,
                IsHitTestVisible = canTouch
            };
            _hitTestElements.Add(hitArea);

            // 2. 【描画用】 実際に見える線
            var line = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = brush,
                StrokeThickness = 2,
                IsHitTestVisible = false // 描画専用（イベントは hitArea で取る）
            };

            // マスク状態なら点線にする
            if (Data.IsMasked)
            {
                line.StrokeDashArray = new DoubleCollection { 2, 2 };
            }

            // --- イベント処理 ---
            // マウスホバーで色を変える
            hitArea.MouseEnter += (s, e) => line.Stroke = SettingsService.Instance.OverBrush;
            hitArea.MouseLeave += (s, e) => line.Stroke = brush;

            // ドラッグ開始
            hitArea.MouseLeftButtonDown += (s, e) => {
                _draggingRect = hitArea;
                _isResizing = true;
                _currentResizeDir = dir;
                // 親の Canvas 基準で座標を取得
                _lastMousePosition = e.GetPosition(this.Parent as UIElement);
                this.CaptureMouse();
                e.Handled = true;
            };

            container.Children.Add(line);    // 下に実線
            container.Children.Add(hitArea); // 上に当たり判定

            return container;
        }

        private void OnFillAreaMouseDown(object sender, MouseButtonEventArgs e)
        {
            // ダブルクリック判定（クロップ範囲の切り替えなど）
            if (e.ClickCount == 2)
            {
                DoubleClicked?.Invoke(this);
                e.Handled = true;
                return;
            }

            _draggingRect = _fillArea;
            _isResizing = false;
            _currentResizeDir = ResizeDirection.Move; // 移動モードをセット

            // 親 Canvas 基準で座標を保持
            _lastMousePosition = e.GetPosition(this.Parent as UIElement);

            CaptureMouse();
            e.Handled = true;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_draggingRect == null) return;

            var parentCanvas = this.Parent as Canvas;
            if (parentCanvas == null) return;

            Point currentPos = e.GetPosition(parentCanvas);
            double diffX = currentPos.X - _lastMousePosition.X;
            double diffY = currentPos.Y - _lastMousePosition.Y;

            double pw = parentCanvas.ActualWidth;
            double ph = parentCanvas.ActualHeight;

            if (_currentResizeDir == ResizeDirection.Move)
            {
                // 💡 座標を比率(0.0〜1.0)に変換してモデルを更新
                Data.X += diffX / pw;
                Data.Y += diffY / ph;
            }
            else if (_isResizing)
            {
                // 💡 選択された辺に応じて計算を切り替え
                switch (_currentResizeDir)
                {
                    case ResizeDirection.Right:
                        Data.Width = Math.Max(0.01, Data.Width + diffX / pw);
                        break;
                    case ResizeDirection.Left:
                        // 左端を動かすときは、位置(X)を動かしつつ幅を逆方向に調整
                        if (Data.Width - diffX / pw > 0.01) { Data.X += diffX / pw; Data.Width -= diffX / pw; }
                        break;
                    case ResizeDirection.Bottom:
                        Data.Height = Math.Max(0.01, Data.Height + diffY / ph);
                        break;
                    case ResizeDirection.Top:
                        // 上端を動かすときは、位置(Y)を動かしつつ高さを逆方向に調整
                        if (Data.Height - diffY / ph > 0.01) { Data.Y += diffY / ph; Data.Height -= diffY / ph; }
                        break;
                }
            }

            // 💡 モデルの変更をUI（座標）に即座に反映
            ApplyDataToUI();

            _lastMousePosition = currentPos;
            e.Handled = true;
        }

        private void EndDrag(object sender, MouseButtonEventArgs e)
        {
            if (IsMouseCaptured)
            {
                ReleaseMouseCapture();
                _draggingRect = null;
                _isResizing = false;
                _currentResizeDir = ResizeDirection.None;
                e.Handled = true;
            }
        }

        private static Brush CreateHatchBrush()
        {
            // 1. 斜線を一本引くための GeometryDrawing を作成
            var lineGeometry = new LineGeometry(new Point(0, 1), new Point(1, 0));
            var lineDrawing = new GeometryDrawing(
                null,
                new System.Windows.Media.Pen(Brushes.Black, 0.2), // 斜線の色と太さ
                lineGeometry
            );

            // 2. DrawingBrush にして、タイル状に並べる設定を行う
            var brush = new DrawingBrush(lineDrawing)
            {
                TileMode = TileMode.Tile,
                Viewport = new Rect(0, 0, 10, 10), // タイル1つのサイズ（10px四方）
                ViewportUnits = BrushMappingMode.Absolute,
                // 背面に薄いグレーを敷くと、より「領域」として認識しやすくなります
                RelativeTransform = new RotateTransform(0)
            };

            // 背景をうっすら黒くしたい場合は、DrawingGroup で合成します
            var group = new DrawingGroup();
            // 下地：20%透明の黒
            group.Children.Add(new GeometryDrawing(new SolidColorBrush(Color.FromArgb(50, 0, 0, 0)), null, new RectangleGeometry(new Rect(0, 0, 1, 1))));
            // 上に斜線
            group.Children.Add(lineDrawing);

            return new DrawingBrush(group)
            {
                TileMode = TileMode.Tile,
                Viewport = new Rect(0, 0, 8, 8),
                ViewportUnits = BrushMappingMode.Absolute
            };
        }
    }
}
