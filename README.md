[![Nuget](https://img.shields.io/nuget/v/Excel2SqlServer)](https://www.nuget.org/packages/Excel2SqlServer/)

This is a library for importing Excel spreadsheets into SQL Server tables using [Excel Data Reader](https://github.com/ExcelDataReader/ExcelDataReader).

Nuget package: **Excel2SqlServer**

In a nutshell, use the [ExcelLoader](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs) class and call one of the `SaveAsync` [overloads](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L108-L130). You can use a local filename or a stream as input. Here's a simple example that loads a single table from a local file:

```csharp
using (var cn = GetConnection())
{
    var loader = new ExcelLoader();
    await loader.SaveAsync("MyFile.xlsx", cn, "dbo", "MyTable");
}
```
This will save an Excel file called `MyFile.xlsx` to a database table `dbo.MyTable`. The table is created if it doesn't exist. Note also there is an `int identity(1,1)` column created called `Id` if it doesn't already exist in the spreadsheet.

If a spreadsheet has multiple sheets and you want to import all the sheets into multiple tables, omit the schema and table name from the `SaveAsync` call. `ExcelLoader` will use the sheet names in the spreadsheet to build the table names. If you need to customize the table names, you can pass a `Dictionary<string, ObjectName>` where the key represents the sheet name, and the `ObjectName` is the schema + object of the resulting table.

By default, data is always appended to existing data. You can pass an optional [Options](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/Options.cs) object to customize the load behavior. For example:
```csharp
using (var stream = await blob.OpenReadAsync())
{
    using (var cn = GetConnection())
    {
        var loader = new ExcelLoader();
        int rows = await loader.SaveAsync(stream, cn, "dbo", "MyTable", new Options() 
        {
            TruncateFirst = true,
            AutoTrimStrings = true,
            RemoveNonPrintingChars = true,
            CustomColumns = new string[]
            {
                "[IsProcessed] bit NOT NULL DEFAULT (0)",
                "[DateUploaded] datetime NOT NULL DEFAULT getdate()"
            }
        });
    }
}
```
This will append some extra columns to the table when it's created `IsProcessed` and `DateUploaded`.

## An encoding error you might see

Note, if you see an error like this...

![img](https://adamosoftware.blob.core.windows.net:443/images/encoding-error.png)

...try adding this line before you use `ExcelLoader`:


```csharp
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
```

## Reference
- Task [CreateTableAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L19)
 (string fileName, SqlConnection connection, string schemaName, string tableName, [ IEnumerable<string> customColumns ])
- Task [CreateTableAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L25)
 (Stream stream, SqlConnection connection, string schemaName, string tableName, [ IEnumerable<string> customColumns ])
- Task\<int\> [SaveAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L109)
 (string fileName, SqlConnection connection, [ Dictionary<string, ObjectName> tableNames ], [ [Options](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/Options.cs#L5) options ])
- Task\<int\> [SaveAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L115)
 (string fileName, SqlConnection connection, string schemaName, string tableName, [ [Options](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/Options.cs#L5) options ])
- Task\<int\> [SaveAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L121)
 (Stream stream, SqlConnection connection, [ Dictionary<string, ObjectName> tableNames ], [ [Options](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/Options.cs#L5) options ])
- Task\<int\> [SaveAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L127)
 (Stream stream, SqlConnection connection, string schemaName, string tableName, [ [Options](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/Options.cs#L5) options ])
- Task\<DataSet\> [ReadAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L241)
 (string fileName)
- Task\<DataSet\> [ReadAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L249)
 (Stream stream)


## Inline Lookup Feature
Use the [InlineLookup\<T\>](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/InlineLookup.cs) class to convert string values to corresponding key values. The `T` generic argument indicates the type of keys being used. Currently `int` is the only type supported. For users who need to upload spreadsheets with key values, allowing them to use text values instead of numeric keys can make an upload process easier.

- See the integration [test](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Testing/LoadTests.cs#L105) showing this in use along with the sample [Excel file](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Testing/Resources/inline-lookup.xlsx) it uses.
    
- Use the [InlineLookup.ExecuteAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/InlineLookup.cs#L31) method to generate a table of mapped key values from a user's upload.
    
- Use the [InlineLookup.GetErrorsAsync](https://github.com/adamfoneil/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/InlineLookup.cs#L98) method to find text values that don't have a mapping.
