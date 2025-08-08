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

                logger.LogInformation("Looking for license at: {Path}", licensePath);

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
                    var searchPath = AppDomain.CurrentDomain.BaseDirectory;
                    logger.LogInformation("Files in application directory: {Files}", 
                        string.Join(", ", Directory.GetFiles(searchPath, "*.lic")));
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
                logger.LogInformation("Starting validation for HTML to Word conversion");
                
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

                logger.LogInformation("Validation passed. Starting conversion from {Input} to {Output}", 
                    report.InputHtmlPath, report.OutputWordPath);

                // Preprocess HTML to clean up problematic formatting
                logger.LogInformation("Preprocessing HTML to clean formatting issues");
                var processedHtmlPath = HtmlPreprocessor.PreprocessHtml(report.InputHtmlPath, logger);

                var loadOptions = new HtmlLoadOptions
                {
                    LoadFormat = LoadFormat.Html,
                    BaseUri = Path.GetDirectoryName(report.InputHtmlPath) ?? string.Empty,
                    WarningCallback = new ConversionWarningCallback(logger)
                };

                logger.LogInformation("Loading preprocessed HTML document with base URI: {BaseUri}", loadOptions.BaseUri);
                var doc = new Document(processedHtmlPath, loadOptions);
                logger.LogInformation("HTML document loaded successfully. Page count: {PageCount}", doc.PageCount);

                // Only update content if user has made CHANGES to the extracted content
                // Don't add if the content is the same as what was extracted
                logger.LogInformation("Checking for user modifications to apply");
                
                // For now, just convert as-is to avoid duplication
                // TODO: Implement proper content replacement logic
                logger.LogInformation("Converting HTML to Word with formatting preservation");
                
                DocumentFormattingHelper.ConfigureHeadersAndFooters(doc, logger);
                DocumentFormattingHelper.EnsureTableHeaderRepetition(doc, logger);
                
                // Prevent orphaned headings at bottom of pages
                DocumentPostProcessor.PreventOrphanedHeadings(doc, logger);

                logger.LogInformation("Saving document to: {Output}", report.OutputWordPath);
                doc.Save(report.OutputWordPath, SaveFormat.Docx);

                // Clean up temporary preprocessed file
                if (processedHtmlPath != report.InputHtmlPath && File.Exists(processedHtmlPath))
                {
                    try
                    {
                        File.Delete(processedHtmlPath);
                        logger.LogDebug("Cleaned up temporary file: {Path}", processedHtmlPath);
                    }
                    catch { /* Ignore cleanup errors */ }
                }

                logger.LogInformation("Conversion completed successfully. File saved to: {Path}", report.OutputWordPath);
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error during document conversion: {Message}", ex.Message);
                return false;
            }
        }

        private void UpdateDocumentContent(Document doc, StatusReport report)
        {
            try
            {
                var builder = new DocumentBuilder(doc);
                bool contentUpdated = false;

                if (!string.IsNullOrWhiteSpace(report.CurrentWeekStatus))
                {
                    var currentWeekNode = FindSectionByHeading(doc, AppConstants.DocumentSections.CURRENT_WEEK_KEYWORDS);
                    if (currentWeekNode != null)
                    {
                        builder.MoveTo(currentWeekNode);
                        if (currentWeekNode.NextSibling != null)
                        {
                            builder.MoveTo(currentWeekNode.NextSibling);
                        }
                        builder.Writeln();
                        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                        builder.Writeln(report.CurrentWeekStatus);
                        logger.LogInformation("Updated current week status");
                        contentUpdated = true;
                    }
                }

                if (!string.IsNullOrWhiteSpace(report.NextWeekGoals))
                {
                    var nextWeekNode = FindSectionByHeading(doc, AppConstants.DocumentSections.NEXT_WEEK_KEYWORDS);
                    if (nextWeekNode != null)
                    {
                        builder.MoveTo(nextWeekNode);
                        if (nextWeekNode.NextSibling != null)
                        {
                            builder.MoveTo(nextWeekNode.NextSibling);
                        }
                        builder.Writeln();
                        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                        builder.Writeln(report.NextWeekGoals);
                        logger.LogInformation("Updated next week goals");
                        contentUpdated = true;
                    }
                }

                if (report.Risks.Any())
                {
                    var risksNode = FindSectionByHeading(doc, AppConstants.DocumentSections.RISK_KEYWORDS);
                    if (risksNode != null)
                    {
                        builder.MoveTo(risksNode);
                        if (risksNode.NextSibling != null)
                        {
                            builder.MoveTo(risksNode.NextSibling);
                        }
                        builder.Writeln();
                        RiskTableBuilder.BuildRiskTable(builder, report);
                        logger.LogInformation("Added {Count} risks to document", report.Risks.Count);
                        contentUpdated = true;
                    }
                }
                
                if (!contentUpdated)
                {
                    logger.LogWarning("No matching sections found in document to update");
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error updating document content");
                throw;
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