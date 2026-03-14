using CommunityToolkit.Mvvm.ComponentModel;
using System.Text.Json.Serialization;
using TraceShot.Controls;

namespace TraceShot.Models
{
    public partial class Bookmark : ObservableObject
    {
        [ObservableProperty] private TimeSpan _time;

        [ObservableProperty] private string? _icon;

        [ObservableProperty] private string? _note;

        [ObservableProperty]  private bool _isListening;

        public string? ImagePath { get; set; }

        public List<AnnotationBase> Annotations { get; set; } = [];

        [JsonIgnore]
        public IEnumerable<RectAnnotation> Rects => Annotations.OfType<RectAnnotation>();

        [JsonIgnore]
        public IEnumerable<NoteAnnotation> Notes => Annotations.OfType<NoteAnnotation>();

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