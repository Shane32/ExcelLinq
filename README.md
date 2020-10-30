# ExcelLinq

[![NuGet](https://img.shields.io/nuget/v/Shane32.ExcelLinq.svg)](https://www.nuget.org/packages/Shane32.ExcelLinq)

ExcelLinq lets you define an Excel workbook much like an Entity Framework data context, where sheets are tables, objects are rows, and properties are columns.  Once the `ExcelContext` is defined, you can load or save a workbook in a single call, add or enumerate data contained within, and use Linq to run any in-memory processing.  The context does not support primary/unique key definitions, nor can/does it enforce relationships across tables (sheets).  There is hooks available to provide basic formatting to the generated Excel file, to parse headers or filter rows from Excel sheets, or to append totals.  Sheets and columns can match based on name when deserializing an Excel workbook.

The project includes extensive testing for its codebase; please see `TestFileContext` for a sample of a context definition including formatting of an Excel workbook, and `EndToEnd.ReadAndWrite` for a sample of reading a file and writing it to another file.
