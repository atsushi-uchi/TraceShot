using CommunityToolkit.Mvvm.ComponentModel;

namespace TraceShot.Models
{
    public partial class Bookmark : ObservableObject
    {
        [ObservableProperty] private TimeSpan _time;

        [ObservableProperty] private string? _icon;

        [ObservableProperty] private string? _note;

        [ObservableProperty]  private bool _isListening;

        public string? ImagePath { get; set; }
        public List<MarkRect> MarkRects { get; set; } = [];

        public List<BalloonNote> Balloons { get; set; } = [];

        public string AddNewLine(string text)
        {
            if (string.IsNullOrEmpty(Note))
            {
                Note = text;
            }
            else
            {
                Note += $"{Environment.NewLine}{text}";
            }
            return Note;
        }

        //public override string ToString()
        //{
        //    // mm: 分 (2桁)
        //    // ss: 秒 (2桁)
        //    // fff: ミリ秒 (3桁)
        //    string timeStr = Time.ToString(@"mm\:ss\.fff");

        //    return $"📌 [{timeStr}] - {Note}";
        //}
    }
}