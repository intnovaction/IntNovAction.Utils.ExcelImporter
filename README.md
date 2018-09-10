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
}

// Read the excel file 
using (var stream = OpenExcel())
{
    // Create the typed importer
    var importer = new Importer<SampleImportInto>();

    // Configure and import
    var importResults = importer
        .SetErrorStrategy(ErrorStrategy.AddElement)
        .FromExcel(stream, "Data With Errors")
        .For(p => p.IntColumn, "Int Column")
        .For(p => p.FloatColumn, "Float Column")
        .For(p => p.DecimalColumn, "Decimal Column")
        .For(p => p.NullableIntColumn, "Nullable Int Column")
        .For(p => p.NullableFloatColumn, "Nullable Float Column")
        .For(p => p.NullableDecimalColumn, "Nullable Decimal Column")
        .For(p => p.StringColumn, "String Column")
        .For(p => p.DateColumn, "Date Column")
        .For(p => p.NullableDateColumn, "Nullable Date Column")
        .Import();
    
    // Read the imported objects...
    var numImportedItems = importResults.ImportedItems.Count();

    // Check the errors...
	var numImportedObjects = importResults.Errors.Count();
}

```
