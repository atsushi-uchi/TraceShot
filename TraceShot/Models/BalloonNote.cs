using WpfPoint = System.Windows.Point; // WPFの座標

public class BalloonNote
{
    public WpfPoint TargetPoint { get; set; } // 💡 始点：赤い丸などで示す場所
    public WpfPoint TextPoint { get; set; }   // 💡 終点：テキストボックスが表示される場所
    public string Text { get; set; } = "";
}