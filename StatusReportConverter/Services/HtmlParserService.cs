using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;
using StatusReportConverter.Constants;
using StatusReportConverter.Models;

namespace StatusReportConverter.Services
{
    public class HtmlParserService : IHtmlParserService
    {
        private readonly ILogger<IHtmlParserService> logger;

        public HtmlParserService(ILogger<IHtmlParserService> logger)
        {
            this.logger = logger;
        }

        public StatusReport ExtractContentFromHtml(string htmlPath)
        {
            var report = new StatusReport { InputHtmlPath = htmlPath };

            try
            {
                if (!File.Exists(htmlPath))
                {
                    logger.LogWarning("HTML file not found: {Path}", htmlPath);
                    return report;
                }

                var doc = new HtmlDocument();
                doc.Load(htmlPath);

                ExtractCurrentWeekStatus(doc, report);
                ExtractNextWeekGoals(doc, report);
                ExtractRisks(doc, report);

                logger.LogInformation("Successfully extracted content from HTML");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error parsing HTML file");
            }

            return report;
        }

        private void ExtractCurrentWeekStatus(HtmlDocument doc, StatusReport report)
        {
            try
            {
                var currentWeekSection = FindSectionByHeading(doc, 
                    AppConstants.DocumentSections.CURRENT_WEEK_KEYWORDS);
                
                if (currentWeekSection != null)
                {
                    var content = ExtractSectionContent(currentWeekSection);
                    report.CurrentWeekStatus = FormatExtractedContent(content);
                    logger.LogInformation("Extracted current week status");
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting current week status");
            }
        }

        private void ExtractNextWeekGoals(HtmlDocument doc, StatusReport report)
        {
            try
            {
                var nextWeekSection = FindSectionByHeading(doc, 
                    AppConstants.DocumentSections.NEXT_WEEK_KEYWORDS);
                
                if (nextWeekSection != null)
                {
                    var content = ExtractSectionContent(nextWeekSection);
                    report.NextWeekGoals = FormatExtractedContent(content);
                    logger.LogInformation("Extracted next week goals");
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting next week goals");
            }
        }

        private void ExtractRisks(HtmlDocument doc, StatusReport report)
        {
            try
            {
                var riskSection = FindSectionByHeading(doc, 
                    AppConstants.DocumentSections.RISK_KEYWORDS);
                
                if (riskSection != null)
                {
                    var table = riskSection.SelectSingleNode(".//following-sibling::table[1]") 
                        ?? riskSection.SelectSingleNode(".//following::table[1]");
                    
                    if (table != null)
                    {
                        ExtractRiskTable(table, report);
                        logger.LogInformation("Extracted {Count} risks", report.Risks.Count);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting risks");
            }
        }

        private HtmlNode? FindSectionByHeading(HtmlDocument doc, string[] keywords)
        {
            var headings = doc.DocumentNode.SelectNodes("//h1 | //h2 | //h3 | //h4");
            
            if (headings != null)
            {
                foreach (var heading in headings)
                {
                    var text = heading.InnerText.ToLower();
                    if (keywords.Any(keyword => text.Contains(keyword)))
                    {
                        return heading;
                    }
                }
            }

            return null;
        }

        private string ExtractSectionContent(HtmlNode sectionHeading)
        {
            var content = new System.Text.StringBuilder();
            var currentNode = sectionHeading.NextSibling;

            while (currentNode != null && !IsHeading(currentNode))
            {
                if (currentNode.Name == "ul" || currentNode.Name == "ol")
                {
                    ExtractListContent(currentNode, content);
                }
                else if (currentNode.Name == "p" || currentNode.Name == "#text")
                {
                    var text = currentNode.InnerText.Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        content.AppendLine(text);
                    }
                }

                currentNode = currentNode.NextSibling;
            }

            return content.ToString();
        }

        private void ExtractListContent(HtmlNode listNode, System.Text.StringBuilder content)
        {
            var items = listNode.SelectNodes(".//li");
            if (items != null)
            {
                foreach (var item in items)
                {
                    var text = item.InnerText.Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        content.AppendLine($"• {text}");
                    }
                }
            }
        }

        private void ExtractRiskTable(HtmlNode table, StatusReport report)
        {
            var rows = table.SelectNodes(".//tr");
            if (rows == null || rows.Count <= 1) return;

            for (int i = 1; i < rows.Count; i++)
            {
                var cells = rows[i].SelectNodes(".//td");
                if (cells != null && cells.Count >= 4)
                {
                    var risk = new Risk
                    {
                        Description = cells[0].InnerText.Trim(),
                        Impact = cells.Count > 1 ? cells[1].InnerText.Trim() : "",
                        Mitigation = cells.Count > 2 ? cells[2].InnerText.Trim() : "",
                        Status = cells.Count > 3 ? cells[3].InnerText.Trim() : "Open",
                        DateIdentified = DateTime.Now
                    };

                    if (!string.IsNullOrWhiteSpace(risk.Description))
                    {
                        report.Risks.Add(risk);
                    }
                }
            }
        }

        private bool IsHeading(HtmlNode node)
        {
            return node.Name == "h1" || node.Name == "h2" || 
                   node.Name == "h3" || node.Name == "h4";
        }

        private string FormatExtractedContent(string content)
        {
            content = Regex.Replace(content, @"[\r\n]+", "\n");
            content = Regex.Replace(content, @"^\s+", "", RegexOptions.Multiline);
            return content.Trim();
        }
    }
}