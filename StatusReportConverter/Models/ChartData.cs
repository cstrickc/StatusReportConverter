using System.Collections.Generic;

namespace StatusReportConverter.Models
{
    public class ChartData
    {
        public string Title { get; set; }
        public ChartType Type { get; set; }
        public List<string> Labels { get; set; }
        public List<ChartDataset> Datasets { get; set; }
        public string Position { get; set; } // Where to insert in document
        
        public ChartData()
        {
            Title = string.Empty;
            Type = ChartType.Pie;
            Labels = new List<string>();
            Datasets = new List<ChartDataset>();
            Position = string.Empty;
        }
    }
    
    public class ChartDataset
    {
        public string Label { get; set; }
        public List<double> Data { get; set; }
        public List<string> BackgroundColors { get; set; }
        
        public ChartDataset()
        {
            Label = string.Empty;
            Data = new List<double>();
            BackgroundColors = new List<string>();
        }
    }
    
    public enum ChartType
    {
        Pie,
        Doughnut,
        Bar,
        Column,
        Line
    }
}