using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;

namespace TraceShot
{
    public partial class SelectionWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point p);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        public IntPtr SelectedHWnd { get; private set; } = IntPtr.Zero;
        public string SelectedTitle { get; private set; } = "";

        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        const int GWL_EXSTYLE = -20;
        const int WS_EX_TRANSPARENT = 0x00000020;

        public SelectionWindow()
        {
            InitializeComponent();

            // 手動配置モード
            this.WindowStartupLocation = WindowStartupLocation.Manual;

            // 全モニターを合計した「仮想スクリーン」のサイズを取得
            this.Left = SystemParameters.VirtualScreenLeft;   // 左サブモニタがあればマイナス値になる
            this.Top = SystemParameters.VirtualScreenTop;
            this.Width = SystemParameters.VirtualScreenWidth;
            this.Height = SystemParameters.VirtualScreenHeight;
        }
        private IntPtr _myHandle = IntPtr.Zero;
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // 💡 ウィンドウの準備が完全に整ったタイミングでハンドルを確定させる
            _myHandle = new System.Windows.Interop.WindowInteropHelper(this).Handle;
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var physPos = System.Windows.Forms.Control.MousePosition;

            // 💡 ステップ1: OSレベルで自分をマウス透過状態にする
            int extendedStyle = GetWindowLong(_myHandle, GWL_EXSTYLE);
            SetWindowLong(_myHandle, GWL_EXSTYLE, extendedStyle | WS_EX_TRANSPARENT);

            // 💡 ステップ2: 背面のウィンドウを取得 (これで自分を突き抜ける)
            IntPtr hWnd = WindowFromPoint(physPos);
            hWnd = GetTopLevelWindow(hWnd);

            // 💡 ステップ3: 透過状態を解除して元に戻す (これがないとクリックできなくなる)
            SetWindowLong(_myHandle, GWL_EXSTYLE, extendedStyle);

            if (hWnd != IntPtr.Zero && hWnd != _myHandle)
            {
                // 🔥 ついにここに来ます！
                if (GetWindowRect(hWnd, out RECT rect))
                {
                    if (GetWindowRect(hWnd, out rect))
                    {
                        // 💡 追加：デスクトップやタスクバーなど、画面全体を覆う特殊なウィンドウを除外
                        // 1. 自分のハンドルなら無視
                        if (hWnd == _myHandle) return;

                        // 2. ウィンドウのクラス名を確認（デスクトップやタスクバーを除外するため）
                        StringBuilder className = new StringBuilder(256);
                        GetClassName(hWnd, className, className.Capacity);
                        string cls = className.ToString();

                        // デスクトップ(Progman/WorkerW)やタスクバー(Shell_TrayWnd)は無視する
                        if (cls == "Progman" || cls == "WorkerW" || cls == "Shell_TrayWnd")
                        {
                            HighlightBorder.Visibility = Visibility.Collapsed;
                            SelectedHWnd = IntPtr.Zero;
                            return;
                        }
                        // 💡 1. 現在のモニターのDPIスケールを取得
                        var dpi = VisualTreeHelper.GetDpi(this);

                        // 💡 2. 物理座標を論理座標（WPF単位）に変換し、かつ
                        // ウィンドウの開始位置（this.Left/Top）を引いて「Canvas内の相対位置」にする
                        // サブモニタが左にある場合、this.Left は負の値なので、引くことで正しくオフセットされます
                        double canvasLeft = (rect.Left / dpi.DpiScaleX) - this.Left;
                        double canvasTop = (rect.Top / dpi.DpiScaleY) - this.Top;
                        double canvasWidth = (rect.Right - rect.Left) / dpi.DpiScaleX;
                        double canvasHeight = (rect.Bottom - rect.Top) / dpi.DpiScaleY;

                        // 💡 3. XAMLのBorder（青い枠）を更新
                        HighlightBorder.Width = Math.Max(0, canvasWidth);
                        HighlightBorder.Height = Math.Max(0, canvasHeight);
                        Canvas.SetLeft(HighlightBorder, canvasLeft);
                        Canvas.SetTop(HighlightBorder, canvasTop);

                        // 💡 4. 表示状態にして、選択中のハンドルを保持
                        HighlightBorder.Visibility = Visibility.Visible;
                        SelectedHWnd = hWnd;
                    }
                }
            }
        }


        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            if (SelectedHWnd != IntPtr.Zero)
            {
                // 💡 ウィンドウタイトルの取得
                int length = GetWindowTextLength(SelectedHWnd);
                if (length > 0)
                {
                    var sb = new System.Text.StringBuilder(length + 1);
                    GetWindowText(SelectedHWnd, sb, sb.Capacity);
                    SelectedTitle = sb.ToString();
                }
                else
                {
                    SelectedTitle = "（タイトルなし）";
                }

                this.DialogResult = true; // 選択確定
                this.Close();
            }
        }

        protected override void OnKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape) this.Close(); // キャンセル
        }

        private IntPtr GetTopLevelWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return IntPtr.Zero;

            // 💡 修正：親を辿るが、デスクトップ（IntPtr.Zero）の手前で止める
            IntPtr current = hWnd;
            while (true)
            {
                IntPtr parent = GetParent(current);
                if (parent == IntPtr.Zero) break;
                current = parent;
            }
            return current;
        }
    }
}