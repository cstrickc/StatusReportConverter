using Aspose.Words;
using Microsoft.Extensions.Logging;

namespace StatusReportConverter.Utils
{
    public class ConversionWarningCallback : IWarningCallback
    {
        private readonly ILogger logger;

        public ConversionWarningCallback(ILogger logger)
        {
            this.logger = logger;
        }

        public void Warning(WarningInfo info)
        {
            switch (info.WarningType)
            {
                case WarningType.UnexpectedContent:
                    logger.LogWarning("Unexpected content warning: {Description}", info.Description);
                    break;
                case WarningType.MinorFormattingLoss:
                    logger.LogDebug("Minor formatting loss: {Description}", info.Description);
                    break;
                case WarningType.MajorFormattingLoss:
                    logger.LogWarning("Major formatting loss: {Description}", info.Description);
                    break;
                default:
                    logger.LogInformation("Conversion warning ({Type}): {Description}", 
                        info.WarningType, info.Description);
                    break;
            }
        }
    }
}