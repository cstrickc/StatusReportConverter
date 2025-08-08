using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;
using StatusReportConverter.Models;

namespace StatusReportConverter.Services
{
    public class ChartService : IChartService
    {
        private readonly ILogger<IChartService> logger;

        public ChartService(ILogger<IChartService> logger)
        {
            this.logger = logger;
        }

        public List<ChartData> ExtractChartDataFromHtml(string htmlPath)
        {
            var charts = new List<ChartData>();
            
            try
            {
                var htmlContent = File.ReadAllText(htmlPath);
                
                // Extract chart data from JavaScript in HTML
                // Looking for Chart.js initialization code
                var statusChart = ExtractChartFromScript(htmlContent, "statusChart", 
                    "Top 25 Initiative Status Distribution", StatusReportConverter.Models.ChartType.Doughnut);
                if (statusChart != null)
                {
                    charts.Add(statusChart);
                }
                
                var valueChart = ExtractChartFromScript(htmlContent, "valueChart", 
                    "Top 25 Value Analysis", StatusReportConverter.Models.ChartType.Bar);
                if (valueChart != null)
                {
                    charts.Add(valueChart);
                }
                
                var categoryChart = ExtractChartFromScript(htmlContent, "categoryChart", 
                    "Top 25 AI Category Analysis", StatusReportConverter.Models.ChartType.Pie);
                if (categoryChart != null)
                {
                    charts.Add(categoryChart);
                }
                
                logger.LogInformation("Extracted {Count} charts from HTML", charts.Count);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting chart data from HTML");
            }
            
            return charts;
        }

        private ChartData? ExtractChartFromScript(string htmlContent, string chartId, string title, StatusReportConverter.Models.ChartType type)
        {
            try
            {
                // Find the chart initialization for this specific chart
                var pattern = $@"getElementById\('{chartId}'\)[\s\S]*?new Chart\([^,]+,\s*(\{{[\s\S]*?\}}\s*\}})";
                var match = Regex.Match(htmlContent, pattern);
                
                if (!match.Success)
                {
                    logger.LogWarning("Could not find chart data for {ChartId}", chartId);
                    return null;
                }
                
                var chartConfig = match.Groups[1].Value;
                
                // Extract labels
                var labelsMatch = Regex.Match(chartConfig, @"labels:\s*\[(.*?)\]");
                var labels = new List<string>();
                if (labelsMatch.Success)
                {
                    var labelsStr = labelsMatch.Groups[1].Value;
                    labels = ExtractStringArray(labelsStr);
                }
                
                // Extract data values
                var dataMatch = Regex.Match(chartConfig, @"data:\s*\[([\d,\s.]+)\]");
                var data = new List<double>();
                if (dataMatch.Success)
                {
                    var dataStr = dataMatch.Groups[1].Value;
                    data = dataStr.Split(',')
                        .Select(s => s.Trim())
                        .Where(s => !string.IsNullOrEmpty(s))
                        .Select(s => double.TryParse(s, out var val) ? val : 0)
                        .ToList();
                }
                
                // Extract colors
                var colorsMatch = Regex.Match(chartConfig, @"backgroundColor:\s*\[(.*?)\]");
                var colors = new List<string>();
                if (colorsMatch.Success)
                {
                    var colorsStr = colorsMatch.Groups[1].Value;
                    colors = ExtractStringArray(colorsStr);
                }
                
                var chartData = new ChartData
                {
                    Title = title,
                    Type = type,
                    Labels = labels,
                    Position = "Visuals",
                    Datasets = new List<ChartDataset>
                    {
                        new ChartDataset
                        {
                            Label = title,
                            Data = data,
                            BackgroundColors = colors
                        }
                    }
                };
                
                logger.LogInformation("Extracted chart: {Title} with {Count} data points", title, data.Count);
                return chartData;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting chart {ChartId}", chartId);
                return null;
            }
        }

        private List<string> ExtractStringArray(string arrayStr)
        {
            var result = new List<string>();
            var matches = Regex.Matches(arrayStr, @"'([^']*)'|""([^""]*)""");
            foreach (Match match in matches)
            {
                var value = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                result.Add(value);
            }
            return result;
        }

        public void InsertChartsIntoDocument(Document doc, List<ChartData> charts)
        {
            if (charts == null || !charts.Any())
            {
                logger.LogInformation("No charts to insert");
                return;
            }
            
            try
            {
                var builder = new DocumentBuilder(doc);
                
                // Find the section where charts should be inserted
                var visualsSection = FindVisualsSection(doc);
                if (visualsSection != null)
                {
                    // Remove any existing content after the visuals heading that might be chart placeholders
                    RemoveChartPlaceholders(doc, visualsSection);
                    
                    // Insert a new paragraph after the visuals heading for our charts
                    var para = (Paragraph)visualsSection;
                    var newPara = new Paragraph(doc);
                    para.ParentNode.InsertAfter(newPara, para);
                    
                    // Move builder to the new paragraph we just created
                    builder.MoveTo(newPara);
                    
                    // Insert charts in the new location
                    foreach (var chartData in charts)
                    {
                        InsertChart(builder, chartData);
                    }
                    
                    logger.LogInformation("Inserted {Count} charts into visuals section", charts.Count);
                }
                else
                {
                    // Insert at end if no visuals section found
                    builder.MoveToDocumentEnd();
                    builder.InsertBreak(BreakType.PageBreak);
                    builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
                    builder.Writeln("Visuals: Top 25 Initiative Analysis");
                    builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                    
                    foreach (var chartData in charts)
                    {
                        InsertChart(builder, chartData);
                    }
                    
                    logger.LogInformation("Inserted {Count} charts at document end", charts.Count);
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error inserting charts into document");
            }
        }

        private void RemoveChartPlaceholders(Document doc, Node visualsSection)
        {
            try
            {
                var nodesToRemove = new List<Node>();
                var currentNode = visualsSection.NextSibling;
                
                // Look for nodes that contain EXACTLY the chart titles (not section headings)
                var chartTitles = new[] { 
                    "Top 25 Initiative Status Distribution",
                    "Top 25 Value Analysis", 
                    "Top 25 AI Category Analysis" 
                };
                
                // Only scan a limited range after the visuals section
                int nodesChecked = 0;
                int maxNodesToCheck = 20; // Don't go too far
                
                while (currentNode != null && nodesChecked < maxNodesToCheck)
                {
                    nodesChecked++;
                    
                    if (currentNode.NodeType == NodeType.Paragraph)
                    {
                        var para = (Paragraph)currentNode;
                        var text = para.GetText().Trim();
                        
                        // Only remove if this EXACTLY matches a chart title
                        foreach (var title in chartTitles)
                        {
                            if (text.Equals(title, StringComparison.OrdinalIgnoreCase))
                            {
                                nodesToRemove.Add(currentNode);
                                logger.LogInformation("Will remove chart placeholder: {Text}", text);
                                break;
                            }
                        }
                    }
                    
                    currentNode = currentNode.NextSibling;
                }
                
                // Remove the identified nodes
                foreach (var node in nodesToRemove)
                {
                    node.Remove();
                }
                
                if (nodesToRemove.Any())
                {
                    logger.LogInformation("Removed {Count} chart placeholder text nodes", nodesToRemove.Count);
                }
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Error removing chart placeholders");
            }
        }
        
        private Node? FindVisualsSection(Document doc)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            
            foreach (Paragraph para in paragraphs)
            {
                var text = para.GetText().ToLower();
                // Look for the visuals section - be more flexible with the search
                if ((text.Contains("visual") && (text.Contains("top 25") || text.Contains("initiative"))) ||
                    (text.Contains("visual") && text.Contains("analysis")))
                {
                    logger.LogInformation("Found visuals section: {Text}", para.GetText().Trim());
                    return para;
                }
            }
            
            logger.LogWarning("Could not find visuals section in document");
            return null;
        }

        private void InsertChart(DocumentBuilder builder, ChartData chartData)
        {
            try
            {
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.ParagraphFormat.KeepWithNext = true; // Prevent heading from being orphaned
                builder.Font.Bold = true;
                builder.Writeln(chartData.Title);
                builder.Font.Bold = false;
                builder.ParagraphFormat.KeepWithNext = false; // Reset for next content
                // No extra line break here to keep chart close to its heading
                
                // Insert the appropriate chart type
                Shape shape;
                double width = 288; // 4 inches - smaller for side-by-side layout
                double height = 216; // 3 inches
                
                // Use slightly larger dimensions for circular charts but not square
                if (chartData.Type == StatusReportConverter.Models.ChartType.Pie || 
                    chartData.Type == StatusReportConverter.Models.ChartType.Doughnut)
                {
                    width = 288; // 4 inches
                    height = 288; // 4 inches - square but smaller
                }
                
                switch (chartData.Type)
                {
                    case StatusReportConverter.Models.ChartType.Pie:
                        shape = builder.InsertChart(Aspose.Words.Drawing.Charts.ChartType.Pie, width, height);
                        ConfigurePieChart(shape.Chart, chartData);
                        break;
                        
                    case StatusReportConverter.Models.ChartType.Doughnut:
                        shape = builder.InsertChart(Aspose.Words.Drawing.Charts.ChartType.Doughnut, width, height);
                        ConfigurePieChart(shape.Chart, chartData); // Similar config as pie
                        break;
                        
                    case StatusReportConverter.Models.ChartType.Bar:
                    case StatusReportConverter.Models.ChartType.Column:
                        shape = builder.InsertChart(Aspose.Words.Drawing.Charts.ChartType.Column, width, height);
                        ConfigureBarChart(shape.Chart, chartData);
                        break;
                        
                    default:
                        shape = builder.InsertChart(Aspose.Words.Drawing.Charts.ChartType.Pie, width, height);
                        ConfigurePieChart(shape.Chart, chartData);
                        break;
                }
                
                // Add dark border to the chart
                shape.Stroke.On = true;
                shape.Stroke.Color = Color.Black;
                shape.Stroke.Weight = 2; // 2pt border
                
                builder.Writeln();
                logger.LogInformation("Inserted {Type} chart: {Title}", chartData.Type, chartData.Title);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error inserting chart: {Title}", chartData.Title);
            }
        }

        private void ConfigurePieChart(Chart chart, ChartData chartData)
        {
            // Don't show title in chart since we already have it as text above
            chart.Title.Show = false;
            
            // Clear existing series
            chart.Series.Clear();
            
            if (chartData.Datasets.Any() && chartData.Datasets[0].Data.Any())
            {
                var dataset = chartData.Datasets[0];
                // Use empty string for series name to avoid duplicate labels
                var series = chart.Series.Add("", 
                    chartData.Labels.ToArray(), 
                    dataset.Data.ToArray());
                
                // Show data labels
                series.HasDataLabels = true;
                series.DataLabels.ShowPercentage = true;
                series.DataLabels.ShowValue = false;
                series.DataLabels.ShowLeaderLines = true;
            }
            
            // Configure legend at bottom for side-by-side layout
            chart.Legend.Position = LegendPosition.Bottom;
        }

        private void ConfigureBarChart(Chart chart, ChartData chartData)
        {
            // Don't show title in chart since we already have it as text above
            chart.Title.Show = false;
            
            // Clear existing series
            chart.Series.Clear();
            
            if (chartData.Datasets.Any() && chartData.Datasets[0].Data.Any())
            {
                var dataset = chartData.Datasets[0];
                // Use empty string for series name to avoid duplicate labels
                var series = chart.Series.Add("", 
                    chartData.Labels.ToArray(), 
                    dataset.Data.ToArray());
                
                // Show data labels
                series.HasDataLabels = true;
                series.DataLabels.ShowValue = true;
            }
            
            // Configure axes
            chart.AxisX.Title.Text = "Categories";
            chart.AxisY.Title.Text = "Values";
            chart.AxisX.Title.Show = true;
            chart.AxisY.Title.Show = true;
            
            // Configure legend at bottom for consistent layout
            chart.Legend.Position = LegendPosition.Bottom;
        }
    }
}