using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace TraceShot
{
    public partial class SelectionWindow : Window
    {
        // --- Win32 API Definitions ---
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point p);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        // 公開プロパティ
        public string SelectedWindowTitle { get; private set; } = "";
        public IntPtr SelectedHWnd { get; private set; } = IntPtr.Zero;

        private DispatcherTimer _timer;

        public SelectionWindow()
        {
            InitializeComponent();

            // ウィンドウがフォーカスを失った（＝他の場所をクリックした）時の処理
            this.Deactivated += OnSelectionWindowDeactivated;

            _timer = new DispatcherTimer(DispatcherPriority.Render);
            _timer.Interval = TimeSpan.FromMilliseconds(50);
            _timer.Tick += OnTimerTick;
            _timer.Start();
        }

        private void OnTimerTick(object? sender, EventArgs e)
        {
            var physPos = System.Windows.Forms.Control.MousePosition;

            // マウス位置のウィンドウ情報を一時的に保持
            var (title, hWnd) = GetWindowInfoAt(physPos);

            // XAMLの表示更新
            TxtScreen.Text = $"画面: {title}";
            TxtHandle.Text = $"hWnd: 0x{hWnd.ToInt64():X}";

            // プロパティを常に最新にしておく（閉じた瞬間の値を確定値とするため）
            SelectedWindowTitle = title;
            SelectedHWnd = hWnd;
            System.Diagnostics.Debug.Print($"title: {SelectedWindowTitle} hWnd: {SelectedHWnd.ToInt64():X}");

            UpdatePosition(physPos);
        }

        private void OnSelectionWindowDeactivated(object? sender, EventArgs e)
        {
            // 他のウィンドウをクリックするなどして、このウィンドウがアクティブでなくなったら閉じる
            _timer.Stop();
            this.Close();
        }

        private (string title, IntPtr hWnd) GetWindowInfoAt(System.Drawing.Point pt)
        {
            IntPtr hWnd = WindowFromPoint(pt);
            if (hWnd == IntPtr.Zero) return ("None", IntPtr.Zero);

            IntPtr parent = GetParent(hWnd);
            while (parent != IntPtr.Zero)
            {
                hWnd = parent;
                parent = GetParent(hWnd);
            }

            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();

            return (string.IsNullOrEmpty(title) ? "(No Title)" : title, hWnd);
        }

        private void UpdatePosition(System.Drawing.Point physPos)
        {
            var screen = System.Windows.Forms.Screen.FromPoint(physPos);
            var source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget == null) return;

            // DPIスケーリングの取得
            double dpiX = source.CompositionTarget.TransformToDevice.M11;
            double dpiY = source.CompositionTarget.TransformToDevice.M22;

            // 現在の物理座標をWPF座標に変換
            double wpfMouseX = physPos.X / dpiX;
            double wpfMouseY = physPos.Y / dpiY;

            // 現在のモニターの有効範囲をWPF座標に変換
            double screenLeft = screen.Bounds.Left / dpiX;
            double screenTop = screen.Bounds.Top / dpiY;
            double screenRight = screen.Bounds.Right / dpiX;
            double screenBottom = screen.Bounds.Bottom / dpiY;

            // マウスからのオフセット距離
            double offset = 15;

            // --- 位置計算（基本は右下） ---
            double targetLeft = wpfMouseX + offset;
            double targetTop = wpfMouseY + offset;

            // --- 境界チェックと反転ロジック ---

            // 1. 右端からはみ出すなら、マウスの左側に表示
            if (targetLeft + this.ActualWidth > screenRight)
            {
                targetLeft = wpfMouseX - this.ActualWidth - offset;
            }
            // 2. 左端からはみ出す（またはマウスを左に移動した際）のガード
            if (targetLeft < screenLeft)
            {
                targetLeft = screenLeft; // 画面端に吸着
            }

            // 3. 下端からはみ出すなら、マウスの上側に表示
            if (targetTop + this.ActualHeight > screenBottom)
            {
                targetTop = wpfMouseY - this.ActualHeight - offset;
            }
            // 4. 上端からはみ出す場合のガード
            if (targetTop < screenTop)
            {
                targetTop = screenTop; // 画面端に吸着
            }

            // 最終的な位置を設定
            this.Left = targetLeft;
            this.Top = targetTop;
        }
    }
}