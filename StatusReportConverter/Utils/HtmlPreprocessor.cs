using System;
using System.IO;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;

namespace StatusReportConverter.Utils
{
    public static class HtmlPreprocessor
    {
        public static string PreprocessHtml(string htmlPath, ILogger logger)
        {
            try
            {
                var doc = new HtmlDocument();
                doc.Load(htmlPath);

                // Remove script tags only (keep style for table formatting)
                RemoveNodes(doc, "//script");
                
                // Only remove problematic CSS properties, not all styles
                RemoveProblematicStyles(doc);
                
                // Clean up special characters minimally
                CleanSpecialCharacters(doc);
                
                // Normalize text direction
                NormalizeTextDirection(doc);
                
                // Don't remove empty elements as they might be needed for spacing

                // Save to temporary file
                var tempPath = Path.Combine(Path.GetTempPath(), $"cleaned_{Path.GetFileName(htmlPath)}");
                doc.Save(tempPath);
                
                logger.LogInformation("HTML preprocessed and saved to: {Path}", tempPath);
                return tempPath;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error preprocessing HTML");
                return htmlPath; // Return original if preprocessing fails
            }
        }

        private static void RemoveNodes(HtmlDocument doc, string xpath)
        {
            var nodes = doc.DocumentNode.SelectNodes(xpath);
            if (nodes != null)
            {
                foreach (var node in nodes)
                {
                    node.Remove();
                }
            }
        }

        private static void RemoveProblematicStyles(HtmlDocument doc)
        {
            var allNodes = doc.DocumentNode.SelectNodes("//*[@style]");
            if (allNodes != null)
            {
                foreach (var node in allNodes)
                {
                    var style = node.GetAttributeValue("style", "");
                    
                    // Only remove CSS that causes vertical text or rotation issues
                    // Keep table formatting, colors, borders, etc.
                    if (style.Contains("writing-mode") || 
                        style.Contains("transform") && style.Contains("rotate") ||
                        style.Contains("text-orientation"))
                    {
                        style = Regex.Replace(style, @"writing-mode\s*:\s*[^;]+;?", "", RegexOptions.IgnoreCase);
                        style = Regex.Replace(style, @"transform\s*:\s*[^;]+;?", "", RegexOptions.IgnoreCase);
                        style = Regex.Replace(style, @"text-orientation\s*:\s*[^;]+;?", "", RegexOptions.IgnoreCase);
                        
                        // Clean up multiple semicolons
                        style = Regex.Replace(style, @";\s*;+", ";", RegexOptions.IgnoreCase);
                        style = style.Trim();
                        
                        if (string.IsNullOrWhiteSpace(style))
                        {
                            node.Attributes.Remove("style");
                        }
                        else
                        {
                            node.SetAttributeValue("style", style);
                        }
                    }
                }
            }
        }

        private static void CleanSpecialCharacters(HtmlDocument doc)
        {
            var textNodes = doc.DocumentNode.SelectNodes("//text()");
            if (textNodes != null)
            {
                foreach (var node in textNodes)
                {
                    if (!string.IsNullOrWhiteSpace(node.InnerText))
                    {
                        var text = node.InnerText;
                        
                        // Remove zero-width characters
                        text = Regex.Replace(text, @"[\u200B-\u200D\uFEFF]", "");
                        
                        // Replace non-breaking spaces with regular spaces
                        text = text.Replace('\u00A0', ' ');
                        
                        // Remove other invisible Unicode characters
                        text = Regex.Replace(text, @"[\u2000-\u200F\u2028-\u202F\u205F-\u206F]", " ");
                        
                        // Normalize multiple spaces to single space
                        text = Regex.Replace(text, @"\s+", " ");
                        
                        if (text != node.InnerText)
                        {
                            node.InnerHtml = HtmlDocument.HtmlEncode(text);
                        }
                    }
                }
            }
        }

        private static void NormalizeTextDirection(HtmlDocument doc)
        {
            // Ensure all elements have LTR text direction
            var body = doc.DocumentNode.SelectSingleNode("//body");
            if (body != null)
            {
                body.SetAttributeValue("dir", "ltr");
                
                // Remove any RTL directions
                var rtlNodes = doc.DocumentNode.SelectNodes("//*[@dir='rtl']");
                if (rtlNodes != null)
                {
                    foreach (var node in rtlNodes)
                    {
                        node.Attributes.Remove("dir");
                    }
                }
            }
        }

        private static void RemoveEmptyElements(HtmlDocument doc)
        {
            bool removedAny;
            do
            {
                removedAny = false;
                var emptyNodes = doc.DocumentNode.SelectNodes("//p[not(node())] | //div[not(node())] | //span[not(node())]");
                if (emptyNodes != null)
                {
                    foreach (var node in emptyNodes)
                    {
                        // Don't remove if it has important attributes
                        if (!node.HasAttributes || 
                            (node.GetAttributeValue("id", "") == "" && 
                             node.GetAttributeValue("class", "") == ""))
                        {
                            node.Remove();
                            removedAny = true;
                        }
                    }
                }
            } while (removedAny);
        }
    }
}