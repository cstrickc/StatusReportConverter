namespace StatusReportConverter.Constants
{
    public static class AppConstants
    {
        public const string APPLICATION_TITLE = "Enterprise AI Status Report Converter";
        public const string REPORT_HEADER_TEXT = "Enterprise AI Status Report";
        
        public static class FileDialogs
        {
            public const string HTML_FILTER = "HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*";
            public const string WORD_FILTER = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";
            public const string DEFAULT_EXTENSION = "docx";
            public const string INPUT_DIALOG_TITLE = "Select HTML Status Report";
            public const string OUTPUT_DIALOG_TITLE = "Save Word Document As";
        }
        
        public static class StatusMessages
        {
            public const string READY = "Ready";
            public const string CONVERTING = "Converting...";
            public const string SUCCESS = "Conversion completed successfully!";
            public const string FAILED = "Conversion failed. Check logs for details.";
            public const string CANCELLED = "Operation cancelled";
            public const string EVALUATION_MODE = "Warning: Running in evaluation mode";
            public const string CURRENT_WEEK_UPDATED = "Current week status updated";
            public const string NEXT_WEEK_UPDATED = "Next week goals updated";
        }
        
        public static class RiskStatus
        {
            public const string OPEN = "Open";
            public const string IN_PROGRESS = "In Progress";
            public const string MITIGATED = "Mitigated";
            public const string CLOSED = "Closed";
        }
        
        public static class Environment
        {
            public const string LOG_LEVEL = "LOG_LEVEL";
            public const string LOG_PATH = "LOG_PATH";
            public const string LICENSE_PATH = "LICENSE_PATH";
            public const string DEFAULT_LOG_PATH = "./logs/";
            public const string DEFAULT_LOG_LEVEL = "Information";
            public const string DEFAULT_LICENSE_PATH = "../Aspose.TotalProductFamily.lic";
        }
        
        public static class Logging
        {
            public const string LOG_FILE_PATTERN = "statusreport-.log";
            public const int RETAINED_FILE_COUNT = 7;
        }
        
        public static class DocumentSections
        {
            public static readonly string[] CURRENT_WEEK_KEYWORDS = { "current week", "this week" };
            public static readonly string[] NEXT_WEEK_KEYWORDS = { "next week", "upcoming" };
            public static readonly string[] RISK_KEYWORDS = { "risk", "risks" };
        }
        
        public static class TableDimensions
        {
            public const double ID_COLUMN_WIDTH = 50;
            public const double DESCRIPTION_COLUMN_WIDTH = 150;
            public const double IMPACT_COLUMN_WIDTH = 100;
            public const double MITIGATION_COLUMN_WIDTH = 150;
            public const double STATUS_COLUMN_WIDTH = 70;
            public const double DATE_COLUMN_WIDTH = 80;
        }
        
        public static class FontSettings
        {
            public const string DEFAULT_FONT = "Arial";
            public const int HEADER_FONT_SIZE = 11;
            public const int FOOTER_FONT_SIZE = 10;
        }
    }
}