namespace TraceShot.Models
{
    public class CaseSummary
    {
        public int CaseId { get; set; }

        public string CaseName => $"No {CaseId}";

        public int StepCount { get; set; }

        public bool IsSuccess { get; set; }

        public TestResult FinalResult { get; set; }

        public string Note { get; set; } = string.Empty;

        public TimeSpan StartTime { get; set; }
        public TimeSpan EndTime { get; set; }
    }
}
