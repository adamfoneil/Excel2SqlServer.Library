[![Nuget](https://img.shields.io/nuget/v/Excel2SqlServer)](https://www.nuget.org/packages/Excel2SqlServer/)

This is a library for importing Excel spreadsheets into SQL Server tables using [Excel Data Reader](https://github.com/ExcelDataReader/ExcelDataReader).

Nuget package: **Excel2SqlServer**

In a nutshell, use the [ExcelLoader](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs) class and call one of the Save overloads [Save string](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L67) or [Save Stream](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L73)

```csharp
using (var cn = GetConnection())
{
    var loader = new ExcelLoader();
    loader.Save("MyFile.xlsx", cn, "dbo", "MyTable");
}
```
This will save an Excel file called `MyFile.xlsx` to a database table `dbo.MyTable`. The table is created if it doesn't exist. Note also there is an `int identity(1,1)` column created called `Id`.

By default, data is always appended to existing data. You can pass an optional [Options](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/Options.cs) object to customize the load behavior. For example:
```csharp
using (var stream = await blob.OpenReadAsync())
{
    using (var cn = GetConnection())
    {
        var loader = new ExcelLoader();
        int rows = loader.Save(stream, cn, "dbo", "MyTable", new Options() 
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
