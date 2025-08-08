using System;
using System.Linq;
using Aspose.Words;
using Microsoft.Extensions.Logging;

namespace StatusReportConverter.Utils
{
    public static class DocumentPostProcessor
    {
        public static void PreventOrphanedHeadings(Document doc, ILogger logger)
        {
            try
            {
                logger.LogInformation("Checking for orphaned headings to prevent page break issues");
                
                var builder = new DocumentBuilder(doc);
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                
                foreach (Paragraph para in paragraphs)
                {
                    // Check if this is a heading (H1, H2, H3, etc.)
                    if (IsHeading(para))
                    {
                        // Check if there's enough content after the heading on the same page
                        var nextNode = para.NextSibling;
                        
                        // If heading is near the end of a page with little content following
                        // Add "Keep with next" property to keep it with following content
                        para.ParagraphFormat.KeepWithNext = true;
                        
                        // Only add page breaks for very specific major sections, not all headings
                        // Commented out for now - too aggressive
                        /*
                        if (IsMajorHeading(para))
                        {
                            // Check if we should add a page break before this heading
                            if (ShouldAddPageBreakBefore(para))
                            {
                                para.ParagraphFormat.PageBreakBefore = true;
                                logger.LogDebug("Added page break before heading: {Text}", 
                                    para.GetText().Trim().Substring(0, Math.Min(50, para.GetText().Trim().Length)));
                            }
                        }
                        */
                        
                        // Ensure headings have proper spacing
                        if (para.ParagraphFormat.SpaceAfter < 6)
                        {
                            para.ParagraphFormat.SpaceAfter = 6;
                        }
                        if (para.ParagraphFormat.SpaceBefore < 12)
                        {
                            para.ParagraphFormat.SpaceBefore = 12;
                        }
                    }
                }
                
                logger.LogInformation("Completed heading optimization for page breaks");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error preventing orphaned headings");
            }
        }
        
        private static bool IsHeading(Paragraph para)
        {
            // Check if paragraph uses a heading style
            var style = para.ParagraphFormat.StyleIdentifier;
            return style == StyleIdentifier.Heading1 ||
                   style == StyleIdentifier.Heading2 ||
                   style == StyleIdentifier.Heading3 ||
                   style == StyleIdentifier.Heading4 ||
                   style == StyleIdentifier.Heading5 ||
                   style == StyleIdentifier.Heading6 ||
                   IsLikelyHeadingByFormat(para);
        }
        
        private static bool IsLikelyHeadingByFormat(Paragraph para)
        {
            // Check if it looks like a heading based on formatting
            var text = para.GetText().Trim();
            
            // Skip if too long to be a heading
            if (text.Length > 200 || text.Length < 2)
                return false;
            
            // Check common heading patterns
            var firstRun = para.Runs.FirstOrDefault() as Run;
            var font = para.ParagraphFormat.Style?.Font ?? firstRun?.Font;
            if (font != null)
            {
                // Headings typically have larger or bold fonts
                if (font.Bold || font.Size >= 14)
                {
                    // Check if it's likely a heading based on content
                    return !text.Contains(".") || // No periods (usually)
                           text.EndsWith(":") ||   // Or ends with colon
                           IsKnownHeadingText(text);
                }
            }
            
            return false;
        }
        
        private static bool IsMajorHeading(Paragraph para)
        {
            var style = para.ParagraphFormat.StyleIdentifier;
            var text = para.GetText().Trim().ToLower();
            
            // H1 and H2 are major headings
            if (style == StyleIdentifier.Heading1 || style == StyleIdentifier.Heading2)
                return true;
            
            // Check for known major section headings
            return IsKnownMajorSection(text);
        }
        
        private static bool ShouldAddPageBreakBefore(Paragraph para)
        {
            // Don't add page break if this is the first paragraph in the document
            if (para.PreviousSibling == null && para.ParentNode.PreviousSibling == null)
                return false;
            
            // Check if there's already a page break
            if (para.ParagraphFormat.PageBreakBefore)
                return false;
            
            var text = para.GetText().Trim().ToLower();
            
            // Always add page break for certain sections
            if (IsKnownMajorSection(text))
            {
                // But not if we're already at the top of a page (check if previous paragraph is very recent)
                var prevPara = para.PreviousSibling as Paragraph;
                if (prevPara != null && string.IsNullOrWhiteSpace(prevPara.GetText()))
                {
                    return false; // Likely already at top of page
                }
                return true;
            }
            
            return false;
        }
        
        private static bool IsKnownHeadingText(string text)
        {
            var lowerText = text.ToLower();
            return lowerText.Contains("week") ||
                   lowerText.Contains("accomplishment") ||
                   lowerText.Contains("planned") ||
                   lowerText.Contains("risk") ||
                   lowerText.Contains("status") ||
                   lowerText.Contains("initiative") ||
                   lowerText.Contains("goal") ||
                   lowerText.Contains("next step");
        }
        
        private static bool IsKnownMajorSection(string text)
        {
            return text.Contains("accomplishment") ||
                   text.Contains("planned") ||
                   text.Contains("risk") ||
                   text.Contains("top 25") ||
                   text.Contains("initiative") ||
                   text.Contains("next week") ||
                   text.Contains("current week");
        }
    }
}