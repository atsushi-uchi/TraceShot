namespace TraceShot.Models
{
    public class EvidenceRect
    {
        // 💡 座標とサイズは 0.0 ～ 1.0 の「正規化座標」で保持します。
        // これにより、ウィンドウサイズが変わっても矩形が正しい位置に追従します。

        /// <summary>
        /// 矩形の左上 X座標 (0.0 = 左端, 1.0 = 右端)
        /// </summary>
        public double X { get; set; }

        /// <summary>
        /// 矩形の左上 Y座標 (0.0 = 上端, 1.0 = 下端)
        /// </summary>
        public double Y { get; set; }

        /// <summary>
        /// 矩形の幅 (0.0 ～ 1.0)
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// 矩形の高さ (0.0 ～ 1.0)
        /// </summary>
        public double Height { get; set; }

        // ⭐ クロップ用かどうかを保持
        public bool IsCropArea { get; set; }

        // コンストラクタ（初期化用）
        public EvidenceRect() { }

        public EvidenceRect(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            IsCropArea = false;
        }
    }
}