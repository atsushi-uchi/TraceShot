using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Text.Json.Serialization;
using TraceShot.Controls;

namespace TraceShot.Models
{
    public enum TestResult { None, OK, NG }

    public partial class Bookmark : ObservableObject
    {
        // JSONに保存され、一生変わらないID
        public Guid Id { get; set; } = Guid.NewGuid();

        public string Name { get; set; }
        public void MarkAsDirty() => IsDirty = true;

        // エクスポート対象かどうか
        public bool IsExportEnabled { get; set; } = true;

        // エクスポート時の並び順
        public int ExportOrder { get; set; } = 0;

        [ObservableProperty] private TimeSpan _time;
        [ObservableProperty] private string? _icon;
        [ObservableProperty] private string? _note;
        [ObservableProperty]  private bool _isListening;
        [JsonIgnore]
        [ObservableProperty] private bool _isDirty = true;

        public string? ImagePath { get; set; }

        public ObservableCollection<AnnotationBase> Annotations { get; set; } = [];

        [JsonIgnore]
        public IEnumerable<RectAnnotation> Rects => Annotations.OfType<RectAnnotation>();

        [JsonIgnore]
        public IEnumerable<NoteAnnotation> Notes => Annotations.OfType<NoteAnnotation>();

        [ObservableProperty] private string _testCaseNo = "";
        [ObservableProperty] private TestResult _result = TestResult.None;
        [ObservableProperty] private bool _isCaseStart = false;

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