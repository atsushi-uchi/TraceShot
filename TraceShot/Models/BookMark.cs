using System.Collections.Generic;
using System.Windows;

public class BookMark
{
    public string? Time { get; set; }
    public double Seconds { get; set; }
    public string? Note { get; set; }
    public string? ImagePath { get; set; }
    public List<Rect> MarkRects { get; set; } = new List<Rect>();

    // 💡 吹き出し（コメント）用のリストを追加
    public List<BalloonNote> Balloons { get; set; } = new List<BalloonNote>();

    public override string ToString() => $"📌 [{Time}] {Note}";
}

