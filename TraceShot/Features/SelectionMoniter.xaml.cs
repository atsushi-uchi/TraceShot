using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using MessageBox = System.Windows.MessageBox;

namespace TraceShot
{
    public partial class SelectionMoniter : Window
    {
        public string SelectedDeviceName { get; private set; }
        public string MoniterName { get; private set; }


        private DispatcherTimer _timer;
        private System.Drawing.Rectangle _currentScreenBounds;

        public SelectionMoniter()
        {
            InitializeComponent();

            // 💡 初期化時は透明のまま、仮想スクリーンの外（または左上）に置いておく
            this.Left = SystemParameters.VirtualScreenLeft;
            this.Top = SystemParameters.VirtualScreenTop;
            this.WindowState = WindowState.Normal;
            // タイマー起動
            _timer = new DispatcherTimer(DispatcherPriority.Render);
            _timer.Interval = TimeSpan.FromMilliseconds(10);
            _timer.Tick += OnTimerTick;
            _timer.Start();
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            var physPos = System.Windows.Forms.Control.MousePosition;
            var screen = System.Windows.Forms.Screen.FromPoint(physPos);

            // 💡 初回起動時、またはモニター移動時
            if (_currentScreenBounds != screen.Bounds)
            {
                // 移動・サイズ変更中は隠す
                this.Opacity = 0;
                _currentScreenBounds = screen.Bounds;

                if (this.WindowState == WindowState.Maximized)
                {
                    this.WindowState = WindowState.Normal;
                }

                var source = PresentationSource.FromVisual(this);
                double dpiX = source?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                double dpiY = source?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;

                // 物理座標を論理座標に変換してジャンプ
                this.Left = screen.Bounds.Left / dpiX;
                this.Top = screen.Bounds.Top / dpiY;

                // 💡 最大化する前にサイズも合わせておくと OS の挙動が安定します
                this.Width = screen.Bounds.Width / dpiX;
                this.Height = screen.Bounds.Height / dpiY;

                this.WindowState = WindowState.Maximized;

                // 💡 描画準備が整ったタイミングで「パッ」と表示
                Dispatcher.BeginInvoke(new Action(() => {
                    this.Opacity = 1;
                }), DispatcherPriority.Render);
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var physPos = System.Windows.Forms.Control.MousePosition;
            var screen = System.Windows.Forms.Screen.FromPoint(physPos);

            // 💡 Windowsが認識しているデバイス名（例: "\\.\DISPLAY2"）を取得
            SelectedDeviceName = screen.DeviceName;
            MoniterName = screen.Primary ? "メインモニター" : "サブモニター";

            DialogResult = true;
            Close();
        }
    }
}
