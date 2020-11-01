using System;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Models;

namespace Shane32.ExcelLinq.Tests.Models
{
    public class TestFileContext : ExcelContext
    {
        public TestFileContext(System.IO.Stream stream) : base(stream) { }
        public TestFileContext(string filename) : base(filename) { }
        public TestFileContext(ExcelPackage excelPackage) : base(excelPackage) { }

        protected override void OnModelCreating(ExcelModelBuilder builder)
        {
            Action<ExcelRange> headerFormatter = range => {
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            };
            Action<ExcelRange> numberFormatter = range => range.Style.Numberformat.Format = "#,##0.00";
            var sheet1 = builder.Sheet<Sheet1>();
            sheet1.Column(x => x.Date)
                .HeaderFormatter(headerFormatter)
                .ColumnFormatter(range => range.Style.Numberformat.Format = "MM/dd/yyyy");
            sheet1.Column(x => x.Quantity)
                .HeaderFormatter(headerFormatter)
                .ColumnFormatter(range => range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
            sheet1.Column(x => x.Description)
                .HeaderFormatter(headerFormatter);
            sheet1.Column(x => x.Amount)
                .HeaderFormatter(headerFormatter)
                .ColumnFormatter(numberFormatter);
            sheet1.Column(x => x.Total)
                .WriteSerializer((range, value) => {
                    var sheet = range.Worksheet;
                    var qtyAddress = sheet.Cells[range.Start.Row, 2].Address;
                    var amountAddress = sheet.Cells[range.Start.Row, 4].Address;
                    range.Formula = $"{qtyAddress}*{amountAddress}";
                })
                .HeaderFormatter(headerFormatter)
                .ColumnFormatter(numberFormatter)
                .WritePolisher(range => {
                    range = range[range.Start.Row + 1, range.Start.Column, range.End.Row + 1, range.End.Column];
                    var totalRange = range.Worksheet.Cells[range.End.Row + 1, range.End.Column];
                    totalRange.Formula = $"SUM({range.Address})";
                    totalRange.Style.Numberformat.Format = "$ #,##0.00";
                    range[totalRange.Start.Row, totalRange.Start.Column - 1].Value = "Total";
                });
            sheet1.Column(x => x.Notes)
                .HeaderFormatter(headerFormatter)
                .Optional();
            sheet1.ReadRangeLocator(worksheet => {
                // skip header rows on this sheet
                for (int i = 1; i <= worksheet.Dimension.Rows; i++) {
                    if (worksheet.Cells[i, 1].Text == "Date") {
                        return worksheet.Cells[i, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                    }
                }
                return null;
            });
            sheet1.SkipEmptyRows();
            sheet1.WriteRangeLocator(worksheet => worksheet.Cells[3, 1]);
            sheet1.WritePolisher((worksheet, range) => {
                worksheet.Calculate();
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++) {
                    var column = worksheet.Column(col);
                    column.AutoFit();
                    column.Width *= 1.2;
                }
                worksheet.Cells[1, 1].Value = "This is a test header";
            });

            var sheet2 = builder.Sheet<Class1>("Sheet2");
            sheet2.Column(x => x.IntColumn);
            sheet2.Column(x => x.FloatColumn);
            sheet2.Column(x => x.DoubleColumn);
            sheet2.Column(x => x.StringColumn);
            sheet2.Column(x => x.BooleanColumn);
            sheet2.Column(x => x.DateTimeColumn)
                .ColumnFormatter(range => range.Style.Numberformat.Format = "MM/dd/yyyy hh:mm AM/PM");
            sheet2.Column(x => x.TimeSpanColumn)
                .ColumnFormatter(range => range.Style.Numberformat.Format = "hh:mm:ss");
            sheet2.Column(x => x.UriColumn);
            sheet2.Column(x => x.GuidColumn);
            sheet2.Column(x => x.NullableIntColumn).Optional();
        }

        protected override object OnReadRow(ExcelRange range, ISheetModel model, IColumnModel[] columnMapping)
        {
            //skip rows where the first column is blank
            if (range.Worksheet.Cells[range.Start.Row, 1].Value == null) return null;

            return base.OnReadRow(range, model, columnMapping);
        }

        protected override void OnWriteFile(ExcelWorkbook workbook)
        {
            base.OnWriteFile(workbook);
            workbook.Calculate();
        }

        public class Sheet1
        {
            public DateTime Date;
            public int Quantity;
            public string Description;
            public decimal Amount;
            public decimal Total;
            public string Notes;
        }
    }
}
