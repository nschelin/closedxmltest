using System;
using ClosedXML.Excel;

namespace closedxmltest
{
    class Program
    {

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

        static void Main(string[] args)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Testing");

                worksheet.Cell(1, 1).Value = "Test Column 1";
                worksheet.Cell(1, 2).Value = "Test Column 2";

                SetColorAndBorder(worksheet.Cell("A1"));
                SetColorAndBorder(worksheet.Cell("B1"));

                worksheet.Column(1).AdjustToContents();
                worksheet.Column(2).AdjustToContents();


                for (var i = 1; i <= 10; i++)
                {
                    var rowId = i + 1;
                    System.Console.WriteLine($"Row: {rowId}");
                    if (rowId % 2 == 0)
                    {
                        worksheet.Cell(rowId, 1).Value = $"Row {rowId}: {1}";
                        worksheet.Cell(rowId, 2).Value = $"Row {rowId}: {2}";
                    }
                    else
                    {
                        worksheet.Cell(rowId, 1).Value = $"Row {rowId}: {1}";
                        worksheet.Cell(rowId, 2).Value = $"Row {rowId}: {2}";

                    }
                }

                var autoFilter = worksheet.RangeUsed().SetAutoFilter();

                autoFilter.Column(1).AddFilter("Row 6: 1").AddFilter("Row 3: 1");

                workbook.SaveAs("MyTest.xlsx");

            }
        }
    }
}
