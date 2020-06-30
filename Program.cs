using System;
using ClosedXML.Excel;

namespace closedxmltest
{
    class Program
    {
        static void Main(string[] args)
        {
            int maxRows = 10;
            int maxColumns = 3;

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Testing");

                worksheet.Cell("A1").Value = "Test Column 1";
                worksheet.Cell("B1").Value = "Test Column 2";
                worksheet.Cell("C1").Value = "Test Column 3";

                SetColorAndBorder(worksheet.Cell("A1"));
                SetColorAndBorder(worksheet.Cell("B1"));
                SetColorAndBorder(worksheet.Cell("C1"));

                worksheet.Column(1).AdjustToContents();
                worksheet.Column(2).AdjustToContents();
                worksheet.Column(3).AdjustToContents();


                for (int i = 1; i <= maxRows; i++)
                {
                    int rowId = i + 1;

                    for (int j = 1; j <= maxColumns; j++)
                    {
                        worksheet.Cell(rowId, j).Value = $"Row {rowId - 1}: Col {j}";
                        worksheet.Cell(rowId, j).Value = $"Row {rowId - 1}: Col {j}";
                    }

                }

                var autoFilter = worksheet.RangeUsed().SetAutoFilter();

                autoFilter.Column(1).AddFilter("Row 6: Col 1").AddFilter("Row 3: Col 1");

                workbook.SaveAs("MyTest.xlsx");

            }
        }



        static void SetColorAndBorder(IXLCell cell)
        {
            cell.Style.Fill.BackgroundColor = XLColor.Lime;

            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            cell.Style.Border.SetTopBorderColor(XLColor.Black);
            cell.Style.Border.SetBottomBorderColor(XLColor.Black);
            cell.Style.Border.SetLeftBorderColor(XLColor.Black);
            cell.Style.Border.SetRightBorderColor(XLColor.Black);
        }
    }
}
