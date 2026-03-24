using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Text.Json.Serialization;
using TraceShot.Controls;

namespace TraceShot.Models
{
    public enum TestResult { None, OK, NG, PEND, SS }

    public partial class Bookmark : ObservableObject
    {
        public Guid Id { get; set; } = Guid.NewGuid();

        [ObservableProperty] private int _caseId = 0;

        [JsonIgnore]
        [ObservableProperty] private string _groupName = "";

        public void Modified() => IsDirty = true;
        public bool IsExportEnabled { get; set; } = true;
        public int ExportOrder { get; set; } = 0;

        [ObservableProperty] private TimeSpan _time;
        [ObservableProperty] private string? _icon;
        [ObservableProperty] private string? _note;
        [ObservableProperty]  private bool _isListening;
        [JsonIgnore]
        [ObservableProperty] private bool _isDirty = true;
        [ObservableProperty] private TestResult _result = TestResult.None;

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