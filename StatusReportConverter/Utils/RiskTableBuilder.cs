using System.Linq;
using Aspose.Words;
using StatusReportConverter.Constants;
using StatusReportConverter.Models;

namespace StatusReportConverter.Utils
{
    public static class RiskTableBuilder
    {
        public static void BuildRiskTable(DocumentBuilder builder, StatusReport report)
        {
            if (!report.Risks.Any())
                return;

            builder.Writeln();
            var table = builder.StartTable();
            
            InsertHeaderRow(builder);
            table.FirstRow.RowFormat.HeadingFormat = true;
            
            foreach (var risk in report.Risks)
            {
                InsertRiskRow(builder, risk);
            }
            
            builder.EndTable();
        }

        private static void InsertHeaderRow(DocumentBuilder builder)
        {
            builder.InsertCell();
            builder.CellFormat.Width = AppConstants.TableDimensions.ID_COLUMN_WIDTH;
            builder.Font.Bold = true;
            builder.Write("ID");
            
            builder.InsertCell();
            builder.CellFormat.Width = AppConstants.TableDimensions.DESCRIPTION_COLUMN_WIDTH;
            builder.Write("Description");
            
            builder.InsertCell();
            builder.CellFormat.Width = AppConstants.TableDimensions.IMPACT_COLUMN_WIDTH;
            builder.Write("Impact");
            
            builder.InsertCell();
            builder.CellFormat.Width = AppConstants.TableDimensions.MITIGATION_COLUMN_WIDTH;
            builder.Write("Mitigation");
            
            builder.InsertCell();
            builder.CellFormat.Width = AppConstants.TableDimensions.STATUS_COLUMN_WIDTH;
            builder.Write("Status");
            
            builder.InsertCell();
            builder.CellFormat.Width = AppConstants.TableDimensions.DATE_COLUMN_WIDTH;
            builder.Write("Date Identified");
            
            builder.EndRow();
        }

        private static void InsertRiskRow(DocumentBuilder builder, Risk risk)
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
    }
}