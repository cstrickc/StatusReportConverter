using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Tables;
using Microsoft.Extensions.Logging;
using StatusReportConverter.Models;

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
                var licensePath = Environment.GetEnvironmentVariable("LICENSE_PATH") ?? "../Aspose.TotalProductFamily.lic";
                
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
                if (!File.Exists(report.InputHtmlPath))
                {
                    logger.LogError("Input HTML file not found: {Path}", report.InputHtmlPath);
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
                ConfigureDocumentFormatting(doc);
                EnsureTableHeaderRepetition(doc);

                var outputDir = Path.GetDirectoryName(report.OutputWordPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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
                    var currentWeekNode = FindSectionByHeading(doc, "current week", "this week");
                    if (currentWeekNode != null)
                    {
                        builder.MoveTo(currentWeekNode);
                        builder.MoveToDocumentEnd();
                        builder.Writeln();
                        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                        builder.Writeln(report.CurrentWeekStatus);
                        logger.LogInformation("Updated current week status");
                    }
                }

                if (!string.IsNullOrWhiteSpace(report.NextWeekGoals))
                {
                    var nextWeekNode = FindSectionByHeading(doc, "next week", "upcoming");
                    if (nextWeekNode != null)
                    {
                        builder.MoveTo(nextWeekNode);
                        builder.MoveToDocumentEnd();
                        builder.Writeln();
                        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                        builder.Writeln(report.NextWeekGoals);
                        logger.LogInformation("Updated next week goals");
                    }
                }

                if (report.Risks.Any())
                {
                    var risksNode = FindSectionByHeading(doc, "risk", "risks");
                    if (risksNode != null)
                    {
                        builder.MoveTo(risksNode);
                        builder.MoveToDocumentEnd();
                        builder.Writeln();
                        
                        var table = builder.StartTable();
                        
                        builder.InsertCell();
                        builder.CellFormat.Width = 50;
                        builder.Font.Bold = true;
                        builder.Write("ID");
                        
                        builder.InsertCell();
                        builder.CellFormat.Width = 150;
                        builder.Write("Description");
                        
                        builder.InsertCell();
                        builder.CellFormat.Width = 100;
                        builder.Write("Impact");
                        
                        builder.InsertCell();
                        builder.CellFormat.Width = 150;
                        builder.Write("Mitigation");
                        
                        builder.InsertCell();
                        builder.CellFormat.Width = 70;
                        builder.Write("Status");
                        
                        builder.InsertCell();
                        builder.CellFormat.Width = 80;
                        builder.Write("Date Identified");
                        
                        builder.EndRow();
                        
                        table.FirstRow.RowFormat.HeadingFormat = true;

                        foreach (var risk in report.Risks)
                        {
                            builder.InsertCell();
                            builder.Font.Bold = false;
                            builder.Write(risk.Id);
                            
                            builder.InsertCell();
                            builder.Write(risk.Description);
                            
                            builder.InsertCell();
                            builder.Write(risk.Impact);
                            
                            builder.InsertCell();
                            builder.Write(risk.Mitigation);
                            
                            builder.InsertCell();
                            builder.Write(risk.Status);
                            
                            builder.InsertCell();
                            builder.Write(risk.DateIdentified.ToShortDateString());
                            
                            builder.EndRow();
                        }
                        
                        builder.EndTable();
                        logger.LogInformation("Added {Count} risks to document", report.Risks.Count);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error updating document content");
            }
        }

        private Node? FindSectionByHeading(Document doc, params string[] keywords)
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

        private void ConfigureDocumentFormatting(Document doc)
        {
            try
            {
                var builder = new DocumentBuilder(doc);

                foreach (Section section in doc.Sections)
                {
                    builder.MoveToSection(doc.Sections.IndexOf(section));
                    builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.Font.Name = "Arial";
                    builder.Font.Size = 11;
                    builder.Writeln("Enterprise AI Status Report");

                    builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    builder.Font.Name = "Arial";
                    builder.Font.Size = 10;
                    
                    builder.Write("Page ");
                    builder.InsertField("PAGE", "");
                    builder.Write(" of ");
                    builder.InsertField("NUMPAGES", "");
                }

                logger.LogInformation("Configured document headers and footers");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error configuring document formatting");
            }
        }

        private void EnsureTableHeaderRepetition(Document doc)
        {
            try
            {
                var tables = doc.GetChildNodes(NodeType.Table, true);
                
                foreach (Table table in tables)
                {
                    if (table.FirstRow != null)
                    {
                        table.FirstRow.RowFormat.HeadingFormat = true;
                    }
                }

                logger.LogInformation("Configured table header repetition for {Count} tables", tables.Count);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error configuring table headers");
            }
        }
    }
}