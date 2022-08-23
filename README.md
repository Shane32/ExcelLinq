# Shane32.ExcelLinq

[![NuGet](https://img.shields.io/nuget/v/Shane32.ExcelLinq.svg)](https://www.nuget.org/packages/Shane32.ExcelLinq)
[![Coverage Status](https://coveralls.io/repos/github/Shane32/ExcelLinq/badge.svg?branch=master)](https://coveralls.io/github/Shane32/ExcelLinq?branch=master)

ExcelLinq lets you define an Excel workbook much like an Entity Framework data context, where sheets are tables,
objects are rows, and properties are columns.  Once the `ExcelContext` is defined, you can load or save a workbook
in a single call, add or enumerate data contained within, and use Linq to run any in-memory processing.
The context does not support primary/unique key definitions, nor can/does it enforce relationships across tables (sheets).
There is hooks available to provide basic formatting to the generated Excel file, to parse headers or filter rows from Excel
sheets, or to append totals.  Sheets and columns can match based on name when deserializing an Excel workbook.

The project includes extensive testing for its codebase; please see `TestFileContext` for a sample of a context definition
including formatting of an Excel workbook, and `EndToEnd.ReadAndWrite` for a sample of reading a file and writing it to another file.

## Basic usage

Below is a sample of how to load an Excel workbook and enumerate the data contained within.

1. Set up your data model which should match the Excel worksheet you want to load.

```csharp
public class Sheet1
{
    public DateTime Date { get; set; }
    public int Quantity { get; set; }
    public string Description { get; set; }
    public decimal Amount { get; set;}
    public decimal Total { get; set; }
    public string Notes { get; set; }
}
```

2. Set up a new context

```csharp
public class TestFileContext : ExcelContext
{
    // in order to read files, you'll need one of these constructors
    public TestFileContext(System.IO.Stream stream) : base(stream) { }
    public TestFileContext(string filename) : base(filename) { }
    public TestFileContext(ExcelPackage excelPackage) : base(excelPackage) { }

    // in order to write new files, you'll need a default constructor
    public TestFileContext() : base() { }

    // define an easy way to access the sheet1 table
    public List<Sheet1> Sheet1 => GetSheet<Sheet1>();

    protected override void OnModelCreating(ExcelModelBuilder builder)
    {
        builder.IgnoreSheetNames();             // ignore sheet names; just take the first sheet
        
        var sheet1 = builder.Sheet<Sheet1>();
        sheet1.Column(x => x.Date);             // define a column for the Date property; look for a column with the name "Date"
        sheet1.Column(x => x.Quantity, "Qty");  // for Quantity, look for a column with the name "Qty"
        sheet1.Column(x => x.Description)       // for Description, look for a column named either "Description" or "Desc"
            .AlternateName("Desc");
        sheet1.Column(x => x.Amount);
        sheet1.Column(x => x.Total);
        sheet1.Column(x => x.Notes)
            .Optional();                        // optional columns are not required to be in the Excel workbook

        sheet1.SkipEmptyRows();                 // skip empty rows when reading
    }
}
```

3. Load the workbook

```csharp
var context = new TestFileContext(filename);
foreach (var row in context.Sheet1) {
    // do something with the row
}
```

4. Save a workbook (new or based on an existing one)

```csharp
var context = new TestFileContext();

// add some data to the context
context.Sheet1.Add(new Sheet1 {
    Date = DateTime.Now,
    Quantity = 1,
    Description = "Test",
    Amount = 1.00m,
    Total = 1.00m,
    Notes = "Test"
});

context.SerializeToFile(filename);
```

## Advanced sample

This sample demonstrates applying formatting and formulas to an Excel workbook.

```csharp
public class TestFileContext : ExcelContext
{
    // in order to read files, you'll need one of these constructors
    public TestFileContext(System.IO.Stream stream) : base(stream) { }
    public TestFileContext(string filename) : base(filename) { }
    public TestFileContext(ExcelPackage excelPackage) : base(excelPackage) { }

    // in order to write new files, you'll need a default constructor
    public TestFileContext() : base() { }

    // define an easy way to access the sheets by name
    public List<Sheet1> Sheet1 => GetSheet<Sheet1>();
    public List<MyClass2> Sheet2 => GetSheet<MyClass2>();
        
    protected override void OnModelCreating(ExcelModelBuilder builder)
    {
        // set up a formatter for the header row
        Action<ExcelRange> headerFormatter = range => {
            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        };
        // set up a formatter for number cells
        Action<ExcelRange> numberFormatter = range => range.Style.Numberformat.Format = "#,##0.00";

        // define a sheet which should be called "Sheet1" in the workbook (name inferred from the class name)
        var sheet1 = builder.Sheet<Sheet1>();

        // define a column called "Date" (it can appear in any order in the sheet; it does not
        // have to be first; but when saving they will save in the order defined)
        sheet1.Column(x => x.Date)
            // apply the header formatter to the header row, which adds a thin line under the header text
            .HeaderFormatter(headerFormatter)
            // apply a date format to cells in this column
            .ColumnFormatter(range => range.Style.Numberformat.Format = "MM/dd/yyyy");
            
        sheet1.Column(x => x.Quantity)
            .HeaderFormatter(headerFormatter)
            // here we are centering the data in this column
            .ColumnFormatter(range => range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);

        sheet1.Column(x => x.Description)
            .HeaderFormatter(headerFormatter);

        sheet1.Column(x => x.Amount)
            .HeaderFormatter(headerFormatter)
            .ColumnFormatter(numberFormatter);   // use our number format defined above

        sheet1.Column(x => x.Total)
            // here instead of writing the value, we write a forumla to these cells that multiplies the quantity times the amount
            .WriteSerializer((range, value) => {
                var sheet = range.Worksheet;
                var qtyAddress = sheet.Cells[range.Start.Row, 2].Address;
                var amountAddress = sheet.Cells[range.Start.Row, 4].Address;
                range.Formula = $"{qtyAddress}*{amountAddress}";
            })
            // we still apply formatting to the header and data cells
            .HeaderFormatter(headerFormatter)
            .ColumnFormatter(numberFormatter)
            // when complete with this column, we are adding a sum at the bottom of the page
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

        // anticipate that there may be a header at the top of the sheet; so determine the range to read
        // from by looking for the first cell in column A that contains the word "Date"
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

        // when writing, we want to skip the first two rows
        sheet1.WriteRangeLocator(worksheet => worksheet.Cells[3, 1]);

        // when finishing writing, we can add additional operations
        sheet1.WritePolisher((worksheet, range) => {
            // recalculate the sheet
            worksheet.Calculate();
            // auto-size the column widths
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++) {
                var column = worksheet.Column(col);
                column.AutoFit();
                column.Width *= 1.2;
            }
            // add a header to the page
            worksheet.Cells[1, 1].Value = "This is a test header";
        });

        // we can also define additional sheets in the same workbook

        // here we define a sheet which should be called "Sheet2" in the workbook
        var sheet2 = builder.Sheet<MyClass2>("Sheet2");
        // todo: define some columns for MyClass2 here
    }

    // we can define additional logic when reading rows
    protected override object OnReadRow(ExcelRange range, ISheetModel model, IColumnModel[] columnMapping)
    {
        //skip rows where the first column is blank
        if (range.Worksheet.Cells[range.Start.Row, 1].Value == null)
            return null;

        return base.OnReadRow(range, model, columnMapping);
    }

    // we can define additional logic when writing the workbook
    protected override void OnWriteFile(ExcelWorkbook workbook)
    {
        base.OnWriteFile(workbook);

        // just a sample; we already calculated the workbook within the write polisher
        workbook.Calculate();
    }
}
```

## Credits

Glory to Jehovah, Lord of Lords and King of Kings, creator of Heaven and Earth, who through his Son Jesus Christ,
has reedemed me to become a child of God. -Shane32
