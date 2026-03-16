using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Text.Json.Serialization;
using TraceShot.Controls;

namespace TraceShot.Models
{
    public partial class Bookmark : ObservableObject
    {
        // JSONに保存され、一生変わらないID
        public Guid Id { get; set; } = Guid.NewGuid();

        public string Name { get; set; }
        [JsonIgnore]
        public bool IsDirty { get; set; } = true;
        public void MarkAsDirty() => IsDirty = true;

        [ObservableProperty] private TimeSpan _time;

        [ObservableProperty] private string? _icon;

        [ObservableProperty] private string? _note;

        [ObservableProperty]  private bool _isListening;

        public string? ImagePath { get; set; }

        public ObservableCollection<AnnotationBase> Annotations { get; set; } = [];

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