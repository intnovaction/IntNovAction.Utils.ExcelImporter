# IntNovAction.Utils.ExcelImporter

Utility to import excel data into POCOs

The usage is quite simple:

```c#
// Define the destination POCO class
class SampleImportInto
{
    public int IntColumn { get; set; }
    public int? NullableIntColumn { get; set; }
    public decimal DecimalColumn { get; set; }
    public decimal? NullableDecimalColumn { get; set; }
    public float FloatColumn { get; set; }
    public float? NullableFloatColumn { get; set; }
    public string StringColumn { get; set; } 
    public DateTime DateColumn { get; set; }
    public DateTime? NullableDateColumn { get; set; }
    public bool BooleanColumn { get; set; }
    public bool? NullableBooleanColumn { get; set; }
}

// Read the excel file 
using (var stream = File.OpenRead("d:\ExcelWithData.xlsx"))
{
    // Create the typed importer
    var importer = new Importer<SampleImportInto>();

    // Configure
    importer
        .SetErrorStrategy(ErrorStrategy.AddElement)
        .FromExcel(stream, "Sheet Name")
        .For(p => p.IntColumn, "Int Column")
        .For(p => p.FloatColumn, "Float Column")
        .For(p => p.DecimalColumn, "Decimal Column")
        .For(p => p.NullableIntColumn, "Nullable Int Column")
        .For(p => p.NullableFloatColumn, "Nullable Float Column")
        .For(p => p.NullableDecimalColumn, "Nullable Decimal Column")
        .For(p => p.StringColumn, "String Column")
        .For(p => p.DateColumn, "Date Column")
        .For(p => p.NullableDateColumn, "Nullable Date Column")
        .For(p => p.BooleanColumn, "Boolean Column")
        .For(p => p.NullableBooleanColumn, "Nullable Boolean Column");

    // Import
    var importResults = importer.Import();
    
    // Read the imported objects...
    var numImportedItems = importResults.ImportedItems.Count();

    // Check the errors...
    var numImportedObjects = importResults.Errors.Count();
}
```

## Error Handling

You can control how errors are handled during import using `SetErrorStrategy`:
- `ErrorStrategy.AddElement`: Adds the row even if it has errors.
- `ErrorStrategy.DoNotAddElement`: Skips rows with errors.

## Nested Objects

Supports nested objects (1-1 relationships). The default constructor must create the nested objects:

```c#
class InnerClass { public int PropInt { get; set; } }
class ClassWithInnerClass {
    public InnerClass Inner { get; set; } = new InnerClass();
    public int TestInt { get; set; }
}

var importer = new Importer<ClassWithInnerClass>()
    .For(p => p.Inner.PropInt, "Inner Column")
    .For(p => p.TestInt, "Base Column");
```

## Excel Template Generation

You can generate an Excel template for your import:

```c#
var importer = new Importer<SampleImportInto>()
    .For(p => p.IntColumn, "Int Column")
    .For(p => p.StringColumn, "String Column");

using (var excelStream = importer.GenerateExcel())
{
    // Save or send the template
}
```

### With Sample Data

```c#
var sampleData = new List<SampleImportInto> { new SampleImportInto { IntColumn = 1, StringColumn = "Test" } };
using (var excelStream = importer.GenerateExcel(sampleData))
{
    // Save or send the template with sample rows
}
```

## Duplicated Column Strategies

Control how duplicated columns are handled with `SetDuplicatedColumnsStrategy`:
- `DuplicatedColumnStrategy.TakeFirst`: Uses the first occurrence.
- `DuplicatedColumnStrategy.TakeLast`: Uses the last occurrence.
- `DuplicatedColumnStrategy.RaiseError`: Fails import if duplicates are found.

## Custom row processing with CustomFor

You can execute custom logic for each imported row using `CustomFor`. This method receives a dictionary with the column values and the destination object, allowing you to process or validate data as needed:

```c#
using (var stream = File.OpenRead("d:\ExcelWithData.xlsx"))
{
    var importer = new Importer<SampleImportInto>()
        .FromExcel(stream, "Sheet Name")
        .For(p => p.IntColumn, "Int Column")
        .CustomFor((Dictionary<string, string> rowValues, SampleImportInto destination) =>
        {
            // Example: log or transform values
            Console.WriteLine($"Row IntColumn value: {rowValues["Int Column"]}");
            // You can also modify 'destination' here
        });

    var importResults = importer.Import();
}
```

You can combine `CustomFor` with duplicated column strategies:

```c#
importer
    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.TakeFirst) // or TakeLast, RaiseError
    .CustomFor((rowValues, destination) =>
    {
        // Will receive the value according to the selected strategy
        Console.WriteLine($"IntColumn: {rowValues["Int Column"]}");
    });
```

## Row Index Mapping

You can map the Excel row index to a property in your POCO:

```c#
class SampleImportInto { public int RowIndex { get; set; } /* ... */ }

importer.SetRowIndex(p => p.RowIndex);
```

## Supported Types

The importer supports the following property types:
- `int`, `int?`
- `decimal`, `decimal?`
- `float`, `float?`
- `string`
- `DateTime`, `DateTime?`
- `bool`, `bool?`

## Dependencies & Compatibility

- Requires [.NET Standard 2.0](https://docs.microsoft.com/en-us/dotnet/standard/net-standard) or [.NET Core 3.1](https://docs.microsoft.com/en-us/dotnet/core/dotnet-core-3-1).
- Uses [ClosedXML](https://github.com/ClosedXML/ClosedXML) for Excel file handling.

---