This is a library for importing Excel spreadsheets into SQL Server tables using [Excel Data Reader](https://github.com/ExcelDataReader/ExcelDataReader).

Nuget package: **Excel2SqlServer**

In a nutshell, use the [ExcelLoader](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs) class and call one of the Save overloads [Save string](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L66) or [Save Stream](https://github.com/adamosoftware/Excel2SqlServer.Library/blob/master/Excel2SqlServer.Library/ExcelLoader.cs#L72)

```
using (var cn = GetConnection())
{
    var loader = new ExcelLoader();
    loader.Save("MyFile.xlsx", cn, "dbo", "MyTable");
}
```
This will save an Excel file called `MyFile.xlsx` to a database table `dbo.MyTable`. The table is created if it doesn't exist.
