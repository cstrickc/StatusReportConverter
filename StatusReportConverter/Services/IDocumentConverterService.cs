using System.Threading.Tasks;
using StatusReportConverter.Models;

namespace StatusReportConverter.Services
{
    public interface IDocumentConverterService
    {
        Task<bool> ConvertHtmlToWordAsync(StatusReport report);
        bool ValidateLicense();
    }
}