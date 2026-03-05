using System.Collections.Generic;
using System.Windows;

public class BookMark
{
    public string? Time { get; set; }
    public double Seconds { get; set; }
    public string? Note { get; set; }
    public string? ImagePath { get; set; }
    public List<MarkRect> MarkRects { get; set; } = new ();

    public List<BalloonNote> Balloons { get; set; } = new ();

    public override string ToString() => $"📌 [{Time}] {Note}";
}

