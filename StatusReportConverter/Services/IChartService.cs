using System.Collections.Generic;
using Aspose.Words;
using StatusReportConverter.Models;

namespace StatusReportConverter.Services
{
    public interface IChartService
    {
        void InsertChartsIntoDocument(Document doc, List<ChartData> charts);
        List<ChartData> ExtractChartDataFromHtml(string htmlPath);
    }
}