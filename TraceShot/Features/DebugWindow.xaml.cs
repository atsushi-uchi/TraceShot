using System.Runtime.InteropServices; // Win32 API呼び出しに必要
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace TraceShot.Features
{
    public partial class DebugWindow : Window
    {
        // --- Win32 API の定義 ---
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point p);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        private DispatcherTimer _timer;

        public DebugWindow()
        {
            InitializeComponent();

            _timer = new DispatcherTimer(DispatcherPriority.Render);
            _timer.Interval = TimeSpan.FromMilliseconds(50); // 負荷軽減のため少し間隔を広げてもOK
            _timer.Tick += OnTimerTick;
            _timer.Start();
        }

        private void OnTimerTick(object? sender, EventArgs e)
        {
            var physPos = System.Windows.Forms.Control.MousePosition;

            // --- 外部ウィンドウタイトルの取得ロジック ---
            string windowTitle = GetWindowTitleAt(physPos);
            TxtScreen.Text = $"画面: {windowTitle}";

            // --- 以下、既存の座標・DPI計算ロジック ---
            var screen = System.Windows.Forms.Screen.FromPoint(physPos);
            var source = PresentationSource.FromVisual(this);

            if (source?.CompositionTarget != null)
            {
                double dpiX = source.CompositionTarget.TransformToDevice.M11;
                double dpiY = source.CompositionTarget.TransformToDevice.M22;

                double wpfMouseX = physPos.X / dpiX;
                double wpfMouseY = physPos.Y / dpiY;
                double screenRight = screen.Bounds.Right / dpiX;
                double screenBottom = screen.Bounds.Bottom / dpiY;

                double offset = 15;
                double targetLeft = wpfMouseX + offset;
                if (targetLeft + this.ActualWidth > screenRight)
                    targetLeft = wpfMouseX - this.ActualWidth - offset;

                double targetTop = wpfMouseY + offset;
                if (targetTop + this.ActualHeight > screenBottom)
                    targetTop = wpfMouseY - this.ActualHeight - offset;

                this.Left = targetLeft;
                this.Top = targetTop;

                TxtPhys.Text = $"物理: {physPos.X}, {physPos.Y}";
                TxtWpf.Text = $"WPF : {this.Left:F0}, {this.Top:F0}";
                TxtDpi.Text = $"DPI : {dpiX * 100}%";
                TxtMon.Text = $"MON : #{Array.IndexOf(System.Windows.Forms.Screen.AllScreens, screen)}";
            }
        }

        /// <summary>
        /// 指定した座標にあるウィンドウのタイトルを取得します
        /// </summary>
        private string GetWindowTitleAt(System.Drawing.Point pt)
        {
            // 1. その地点にあるウィンドウハンドルを取得
            IntPtr hWnd = WindowFromPoint(pt);
            if (hWnd == IntPtr.Zero) return "None";

            // 2. 子コントロール（ボタンなど）を掴んだ場合、親ウィンドウを辿る
            // ※タイトルバーを持つウィンドウを取得するため
            IntPtr parent = GetParent(hWnd);
            while (parent != IntPtr.Zero)
            {
                hWnd = parent;
                parent = GetParent(hWnd);
            }

            // 3. タイトルを取得
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, sb.Capacity);

            string title = sb.ToString();
            return string.IsNullOrEmpty(title) ? "(No Title)" : title;
        }
    }
}