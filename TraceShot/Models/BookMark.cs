namespace TraceShot.Models
{
    public class BookMark
    {
        public TimeSpan Time { get; set; }
        public string? Note { get; set; }
        public string? ImagePath { get; set; }
        public List<MarkRect> MarkRects { get; set; } = new();

        public List<BalloonNote> Balloons { get; set; } = new();

        public override string ToString()
        {
            // mm: 分 (2桁)
            // ss: 秒 (2桁)
            // fff: ミリ秒 (3桁)
            string timeStr = Time.ToString(@"mm\:ss\.fff");

            return $"📌 [{timeStr}] - {Note}";
        }
    }
}