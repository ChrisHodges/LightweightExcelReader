# LightweightExcelReader

[![NuGet](https://img.shields.io/nuget/v/LightweightExcelReader.svg)](https://www.nuget.org/packages/LightweightExcelReader)

The fastest way to read Excel data in .NET

### What is this?
It's a [.NET Standard](https://docs.microsoft.com/en-us/dotnet/standard/net-standard) library for reading data from an Excel [.xslx format](https://stackoverflow.com/questions/18334314/what-do-excel-xml-cell-attribute-values-mean/18346273#18346273) spreadsheet. It's compatible with .Net Framework 4.6.1+ and .Net Standard 2.0+.

### How do I get it?
Via Nuget, either the Package Manager Console:

```
PM> Install-Package LightweightExcelReader
```
or the .NET CLI:

```
> dotnet add package LightweightExcelReader
```
### How do I use it?

Like this:

```C#
//Instatiate a spreadsheet reader by file path:
var excelReader = new ExcelReader("/path/to/workbook.xlsx");
//or from a stream:
var excelReaderFromStream = new ExcelReader(stream);

//Get a worksheet by index:
var sheetReader = excelReader[0];
//or by name:
var sheetNamedMyAwesomeSpreadsheet = excelReader["MyAwesomeSpreadsheet"];

//Use ReadNext() to read the next cell in the spreadsheet:
while (sheetReader.ReadNext()) {
    dictionaryOfCellValues.Add(sheetReader.Address, sheetReader.Value);
}
//ReadNext() returns false if the reader has read to the end of the spreadsheet

//Use ReadNextInRow() to read the next cell in the current row:
var dictionaryOfCellValues = new Dictionary<string, object>();
while (sheetReader.ReadNextInRow()) {
    dictionaryOfCellValues.Add(sheetReader.Address, sheetReader.Value);
}
//ReadNextInRow() returns if the reader has read to the end of the current row:

//Get data for a specific cell:
object cellA1Value = sheetReader["A1"];

//For a range:
IEnumerable<object> cellsFromA1ToD4 = sheetReader["A1","D4"];

//for a row:
IEnumerable<object> row3 = sheetReader.Row(3);

//orr a column:
IEnumerable<object> columnB = sheetReader.Column("B");

```
### Is it fast?
You bet. We've aimed to create the fastest Excel reader for .Net, and we think we've succeeded. Included in the repo is a [Benchmark Dot Net](https://github.com/dotnet/BenchmarkDotNet) benchmarking that compares the performance of  **LightweightExcelReader** to [OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/), [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) and a good old fashioned [OleDbDataAdapter](https://docs.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbdataadapter?view=netframework-4.8). The benchmark reads a 500 row spreadsheet to memory. On a 2.7 GHz Quad-Core Intel Core i7 laptop running Windows 10, the results looked like this:

| Package                  |  Operation |
|--------------------------|------------|
|  OpenXml                 | 8.109 ms   |
|  ExcelDataReader         | 6.225 ms   |
|  OleDbDataAdapter        | 52.833 ms  |
| *LightweightExcelReader* | *2.670 ms* |

### Is LightweightExcelReader right for me?
LightweightExcelReader is right for you if:

* You only want to read Excel data - you don't want to create or edit a spreadsheet.
* You only need to read .xlsx format files (Excel 2007 and later).
* You don't want to convert a spreadsheet into an array of objects.


### Support
**LightweightExcelReader** is currently in beta. If you discover a bug please [raise an issue](https://github.com/ChrisHodges/LightweightExcelReader/issues/new) or better still [fork the project](https://help.github.com/en/github/getting-started-with-github/fork-a-repo) and raise a Pull Request if you can fix the issue yourself.
