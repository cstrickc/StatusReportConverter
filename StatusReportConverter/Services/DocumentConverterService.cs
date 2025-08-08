using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Microsoft.Extensions.Logging;
using StatusReportConverter.Constants;
using StatusReportConverter.Models;
using StatusReportConverter.Utils;

namespace StatusReportConverter.Services
{
    public class DocumentConverterService : IDocumentConverterService
    {
        private readonly ILogger<IDocumentConverterService> logger;
        private bool licenseLoaded = false;

        public DocumentConverterService(ILogger<IDocumentConverterService> logger)
        {
            this.logger = logger;
            LoadLicense();
        }

        private void LoadLicense()
        {
            try
            {
                var licensePath = Environment.GetEnvironmentVariable(AppConstants.Environment.LICENSE_PATH) 
                    ?? AppConstants.Environment.DEFAULT_LICENSE_PATH;
                
                if (!Path.IsPathRooted(licensePath))
                {
                    licensePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, licensePath);
                }

                if (File.Exists(licensePath))
                {
                    var license = new License();
                    license.SetLicense(licensePath);
                    licenseLoaded = true;
                    logger.LogInformation("Aspose license loaded successfully from {Path}", licensePath);
                }
                else
                {
                    logger.LogWarning("License file not found at {Path}. Running in evaluation mode.", licensePath);
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to load Aspose license");
            }
        }

        public bool ValidateLicense()
        {
            return licenseLoaded;
        }

        public async Task<bool> ConvertHtmlToWordAsync(StatusReport report)
        {
            return await Task.Run(() => ConvertHtmlToWord(report));
        }

        private bool ConvertHtmlToWord(StatusReport report)
        {
            try
            {
                if (!ValidationHelper.ValidateHtmlFile(report.InputHtmlPath, out var inputError))
                {
                    logger.LogError("Input validation failed: {Error}", inputError);
                    return false;
                }

                if (!ValidationHelper.ValidateOutputPath(report.OutputWordPath, out var outputError))
                {
                    logger.LogError("Output validation failed: {Error}", outputError);
                    return false;
                }

                logger.LogInformation("Starting conversion from {Input} to {Output}", 
                    report.InputHtmlPath, report.OutputWordPath);

                var loadOptions = new HtmlLoadOptions
                {
                    LoadFormat = LoadFormat.Html,
                    BaseUri = Path.GetDirectoryName(report.InputHtmlPath) ?? string.Empty
                };

                var doc = new Document(report.InputHtmlPath, loadOptions);

                UpdateDocumentContent(doc, report);
                DocumentFormattingHelper.ConfigureHeadersAndFooters(doc, logger);
                DocumentFormattingHelper.EnsureTableHeaderRepetition(doc, logger);

                doc.Save(report.OutputWordPath, SaveFormat.Docx);

                logger.LogInformation("Conversion completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error during document conversion");
                return false;
            }
        }

        private void UpdateDocumentContent(Document doc, StatusReport report)
        {
            try
            {
                var builder = new DocumentBuilder(doc);

                if (!string.IsNullOrWhiteSpace(report.CurrentWeekStatus))
                {
                    var currentWeekNode = FindSectionByHeading(doc, AppConstants.DocumentSections.CURRENT_WEEK_KEYWORDS);
                    if (currentWeekNode != null)
                    {
                        builder.MoveTo(currentWeekNode);
                        builder.MoveToDocumentEnd();
                        builder.Writeln();
                        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                        builder.Writeln(ValidationHelper.SanitizeText(report.CurrentWeekStatus));
                        logger.LogInformation("Updated current week status");
                    }
                }

                if (!string.IsNullOrWhiteSpace(report.NextWeekGoals))
                {
                    var nextWeekNode = FindSectionByHeading(doc, AppConstants.DocumentSections.NEXT_WEEK_KEYWORDS);
                    if (nextWeekNode != null)
                    {
                        builder.MoveTo(nextWeekNode);
                        builder.MoveToDocumentEnd();
                        builder.Writeln();
                        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                        builder.Writeln(ValidationHelper.SanitizeText(report.NextWeekGoals));
                        logger.LogInformation("Updated next week goals");
                    }
                }

                if (report.Risks.Any())
                {
                    var risksNode = FindSectionByHeading(doc, AppConstants.DocumentSections.RISK_KEYWORDS);
                    if (risksNode != null)
                    {
                        builder.MoveTo(risksNode);
                        builder.MoveToDocumentEnd();
                        RiskTableBuilder.BuildRiskTable(builder, report);
                        logger.LogInformation("Added {Count} risks to document", report.Risks.Count);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error updating document content");
            }
        }

        private Node? FindSectionByHeading(Document doc, string[] keywords)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            
            foreach (Paragraph para in paragraphs)
            {
                var text = para.GetText().ToLower();
                if (keywords.Any(keyword => text.Contains(keyword)))
                {
                    return para;
                }
            }
            
            return null;
        }
    }
}