using System.Windows;

public class RecordingEvidence
{
    public string? VideoFileName { get; set; }
    public string? RecordingDate { get; set; }
    public string? Mode { get; set; } // 画面全体、矩形、ウィンドウなど
    public List<BookMark> Bookmarks { get; set; } = new List<BookMark>();
}
