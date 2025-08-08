using StatusReportConverter.Models;

namespace StatusReportConverter.Services
{
    public interface IHtmlParserService
    {
        StatusReport ExtractContentFromHtml(string htmlPath);
    }
}