using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Microsoft.Extensions.Logging;
using StatusReportConverter.Constants;

namespace StatusReportConverter.Utils
{
    public static class DocumentFormattingHelper
    {
        public static void ConfigureHeadersAndFooters(Document doc, ILogger logger)
        {
            try
            {
                var builder = new DocumentBuilder(doc);

                foreach (Section section in doc.Sections)
                {
                    builder.MoveToSection(doc.Sections.IndexOf(section));
                    
                    builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.Font.Name = AppConstants.FontSettings.DEFAULT_FONT;
                    builder.Font.Size = AppConstants.FontSettings.HEADER_FONT_SIZE;
                    builder.Writeln(AppConstants.REPORT_HEADER_TEXT);

                    builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    builder.Font.Name = AppConstants.FontSettings.DEFAULT_FONT;
                    builder.Font.Size = AppConstants.FontSettings.FOOTER_FONT_SIZE;
                    
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

        public static void EnsureTableHeaderRepetition(Document doc, ILogger logger)
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