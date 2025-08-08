using System.Collections.ObjectModel;

namespace StatusReportConverter.Models
{
    public class StatusReport
    {
        public string InputHtmlPath { get; set; }
        public string OutputWordPath { get; set; }
        public string CurrentWeekStatus { get; set; }
        public string NextWeekGoals { get; set; }
        public ObservableCollection<Risk> Risks { get; set; }

        public StatusReport()
        {
            InputHtmlPath = string.Empty;
            OutputWordPath = string.Empty;
            CurrentWeekStatus = string.Empty;
            NextWeekGoals = string.Empty;
            Risks = new ObservableCollection<Risk>();
        }
    }
}