namespace TraceShot.Models
{
    public class RecordingEvidence
    {
        public string? VideoFileName { get; set; }
        public string? RecordingDate { get; set; }
        public string? Mode { get; set; } // 全画面、矩形、ウィンドウ
        public List<BookMark> Bookmarks { get; set; } = new ();
    }
}