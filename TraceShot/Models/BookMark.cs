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
        public List<EvidenceRect> Regions { get; set; } = [];

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
    }
}