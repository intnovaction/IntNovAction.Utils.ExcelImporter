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
using (var stream = File.OpenRead("d:\\Exce�WithData.xlsx"))
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

It supports nested objects also (1-1 relationships), but the default constructor must create the nested objects.

It also allows to generate an Excel file with the exact format to import the data, very useful as a template

```c#
	var importer = new Importer<SampleImportInto>();

    // Configure
    importer
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

	using (var excelStream = importer.GenerateExcel())
    {
		// Do stuff with the excel
	}

```