[![GitHub Actions](https://github.com/simplify9/SW-ExportToExcel/actions/workflows/nuget-publish.yml/badge.svg)](https://github.com/simplify9/SW-ExportToExcel/actions/workflows/nuget-publish.yml)
[![NuGet](https://img.shields.io/nuget/v/SimplyWorks.ExportToExcel?style=for-the-badge)](https://www.nuget.org/packages/SimplyWorks.ExportToExcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](LICENSE)

# SimplyWorks.ExportToExcel

Extension methods that write any `IEnumerable<T>` (or `IAsyncEnumerable<T>`) to an `.xlsx` file.

- Rows are streamed straight into the file, so memory stays flat no matter how many rows you export.
- Targets .NET 10 and .NET 8.

```bash
dotnet add package SimplyWorks.ExportToExcel
```

## Quick start

```csharp
using SW.ExportToExcel;

using var stream = new MemoryStream();
await orders.WriteExcelAsync(stream);
```

This writes every readable property of a supported type, in declaration order, with the property name as the header. To choose the columns, rename them or put them in a different order, pass a map from property name to header:

```csharp
await orders.WriteExcelAsync(stream, new Dictionary<string, string>
{
    ["Total"] = "Total (€)",
    ["Customer"] = "Customer",
    ["PlacedOn"] = "Date",
});
```

Columns come out in the order you list them. Names that don't match a property are skipped.

### Computed columns

```csharp
await orders.WriteExcelAsync(stream, new[]
{
    new ExcelColumn<Order>("Customer", o => o.Customer.Name),
    new ExcelColumn<Order>("City", o => o.Customer.Address.City),
    new ExcelColumn<Order>("Total incl. VAT", o => o.Total * 1.2m) { Width = 18 },
});
```

### Dictionary and dynamic rows

Rows of type `IDictionary<string, object>` (including `ExpandoObject`) or `IDictionary<string, string>` are exported by key.

- **Headers:** the keys found in the first rows that are read ahead (100 by default; see `AutoFitSampleRows` below).
- **Later rows:** a key a row doesn't have leaves its cell empty.
- **Choosing columns:** pass a column map to set which keys appear and in what order.

```csharp
var rows = new List<Dictionary<string, object>>
{
    new() { ["Sku"] = "A-1", ["Qty"] = 3 },
    new() { ["Sku"] = "B-2", ["Qty"] = 1, ["Note"] = "fragile" },
};
await rows.WriteExcelAsync(stream);
```

### Streaming from a database

Pass an `IAsyncEnumerable<T>`, such as EF Core's `AsAsyncEnumerable()`, and rows are written as they arrive, without loading the whole result first:

```csharp
await db.Orders.Where(o => o.Year == 2026).AsAsyncEnumerable()
    .WriteExcelAsync(stream, cancellationToken: ct);
```

## Cell types

| Value | Excel cell |
| --- | --- |
| `int`, `long`, `decimal`, `double` and the other numeric types | number |
| `bool` | `TRUE` / `FALSE` |
| `DateTime`, `DateTimeOffset` | date and time (a `DateTimeOffset` keeps its own clock time) |
| `DateOnly` | date |
| `string`, `char`, `Guid`, enums (by name), `TimeSpan`, `TimeOnly` | text |
| `null`, empty string, dates before year 100 | empty cell |
| `NaN`, `Infinity` | the text `NaN` / `Infinity` |

Properties of any other type, such as nested objects and collections, are skipped. Computed columns can return anything; values that aren't listed above are written as text.

Text is cleaned so that it always produces a file Excel can open: characters that XML can't carry (control characters, unpaired surrogates) are removed, and text over Excel's 32,767-character cell limit is truncated.

## Options

```csharp
await orders.WriteExcelAsync(stream, new ExcelExportOptions
{
    SheetName = "Orders",              // invalid characters become '_', truncated to 31
    FreezeHeaderRow = true,            // default true
    BoldHeader = true,                 // default true
    AutoFilter = true,                 // filter drop-downs on the header, default false
    AutoFitColumns = true,             // size columns from the header and the first rows, default true
    AutoFitSampleRows = 100,           // rows read ahead for sizing and dictionary keys
    MaxColumnWidth = 60,
    DateTimeFormat = "yyyy-mm-dd hh:mm", // default: Excel's built-in short date and time
    DateFormat = "yyyy-mm-dd",         // for DateOnly; default: Excel's built-in short date
});
```

## Writing to a web response

The file is written synchronously. ASP.NET Core rejects synchronous writes to a response body, so write to a `MemoryStream` (or a file) first, then return that:

```csharp
using var stream = new MemoryStream();
await rows.WriteExcelAsync(stream);
return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "export.xlsx");
```

## The original API

`WriteExcel(...)`, `ExportToExcel(...)` and `GenerateExcel(...)` are unchanged and produce the same sheets as before:

- Columns always follow the class's property declaration order. A column map only selects and renames columns; it doesn't reorder them.
- Booleans are written as the text `true` / `false`.
- Only `string`, `bool`, `DateTime` and numeric properties are exported.
- Every cell is written, including empty ones.

The original methods run on the same streaming writer, so they get the same speed and memory gains and the same fixes:

- Characters XML can't carry, `NaN` and `Infinity`, over-long text, and types with nothing to export no longer break the export or the file.
- The header row is now actually frozen; in earlier versions the freeze setting was misplaced and Excel ignored it.
- The first column is no longer forced to a width of 4.
- Dictionary rows are matched by key instead of by position.
- An empty dictionary list gives an empty sheet instead of throwing.

To get ordered columns and typed cells, switch to `WriteExcelAsync`. The two differ only as described above.

## Limits

An Excel sheet holds at most 1,048,576 rows and 16,384 columns. Larger exports are still written, but Excel won't open them; other readers may.

## Upgrading from 8.x / 2.x

- **Frameworks:** the package now targets `net10.0` and `net8.0`. Projects on older frameworks can stay on 2.0.5 (`netstandard2.0`).
- **OpenXml 3:** the library no longer uses `DocumentFormat.OpenXml` itself, but still depends on it, now at 3.x instead of 2.13, so projects that get it through this package keep compiling. If you use OpenXml's API directly, check its [3.0 breaking changes](https://github.com/dotnet/Open-XML-SDK/blob/main/CHANGELOG.md).
- **Output changes:** the original methods' output changes only where it was broken before, as listed under [The original API](#the-original-api).

## Getting support 👷

If you encounter any bugs, don't hesitate to submit an [issue](https://github.com/simplify9/SW-ExportToExcel/issues). We'll get back to you promptly!
