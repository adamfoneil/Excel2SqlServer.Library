using AO.Models;
using Dapper;
using Excel2SqlServer.Library;
using Microsoft.Data.SqlClient;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SqlServer.LocalDb;
using SqlServer.LocalDb.Extensions;
using SqlServer.LocalDb.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace Testing
{
    [TestClass]
    public class LoadTests
    {
        private const string dbName = "ExcelImport";

        [TestMethod]
        public void Case01()
        {
            using (var cn = LocalDb.GetConnection(dbName))
            {
                CreateSchema(cn);

                var loader = new ExcelLoader();
                using (var stream = GetResource("case01.xlsx"))
                {
                    loader.SaveAsync(stream, cn, "upload", "Case01").Wait();
                }
            }
        }

        [TestMethod]
        public void Case02()
        {
            using (var cn = LocalDb.GetConnection(dbName))
            {
                CreateSchema(cn);

                using (var stream = GetResource("case02.xlsx"))
                {
                    var loader = new ExcelLoader();
                    loader.SaveAsync(stream, cn, "upload", "Case02", new Options()
                    {
                        TruncateFirst = true,
                        AutoTrimStrings = true,
                        RemoveNonPrintingChars = true
                    }).Wait();
                }

                var lastNames = cn.Query<string>("SELECT [Last Name] FROM [upload].[Case02]");
                Assert.IsTrue(lastNames.All(name => name.Equals(name.Trim())));

                var addresses = cn.Query<string>("SELECT [Address] FROM [upload].[Case02]");
                Assert.IsTrue(addresses.All(address => address.Equals(address.Trim())));
            }
        }

        [TestMethod]
        public void Case03()
        {
            using (var cn = LocalDb.GetConnection(dbName))
            {
                CreateSchema(cn);

                using (var stream = GetResource("case03.xlsx"))
                {
                    new ExcelLoader().SaveAsync(stream, cn).Wait();

                    Assert.IsTrue(cn.Query<int>("SELECT COUNT(1) FROM [dbo].[Companies]").Count() > 0);
                    Assert.IsTrue(cn.Query<int>("SELECT COUNT(1) FROM [dbo].[People]").Count() > 0);
                }
            }
        }

        [TestMethod]
        public void Case03_CustomNames()
        {
            using (var cn = LocalDb.GetConnection(dbName))
            {
                ExecuteIgnoreError(cn, "DROP TABLE [loader].[Company]");
                ExecuteIgnoreError(cn, "DROP TABLE [loader].[Person]");
                ExecuteIgnoreError(cn, "DROP SCHEMA [loader]");

                using (var stream = GetResource("case03.xlsx"))
                {
                    new ExcelLoader().SaveAsync(stream, cn, new Dictionary<string, ObjectName>()
                    {
                        { "Companies", new ObjectName("loader", "Company") },
                        { "People", new ObjectName("loader", "Person") }
                    }).Wait();

                    Assert.IsTrue(cn.Query<int>("SELECT COUNT(1) FROM [loader].[Company]").Count() > 0);
                    Assert.IsTrue(cn.Query<int>("SELECT COUNT(1) FROM [loader].[Person]").Count() > 0);
                }
            }
        }

        [TestMethod]
        public void InlineLookup()
        {
            using (var cn = LocalDb.GetConnection(dbName, SampleLookupObjects()))
            {
                CreateRows(cn, "dbo.Region", "North", "South", "East", "West");
                CreateRows(cn, "dbo.Type", "Flavorful", "Oblong", "Interminable", "Gingersnap");

                var loader = new ExcelLoader();
                using (var xls = GetResource("inline-lookup.xlsx"))
                {
                    loader.SaveAsync(xls, cn, "dbo", "SalesDataRaw", new Options()
                    {
                        TruncateFirst = true
                    }).Wait();

                    var inlineLookup = new InlineLookup<int>("dbo.SalesDataRaw", "Id", "dbo.SalesDataKeyed", new Lookup[]
                    {
                        new Lookup("Region", "RegionId", "dbo.Region", "Name", "Id"),
                        new Lookup("Type", "TypeId", "dbo.Type", "Name", "Id")
                    });

                    inlineLookup.ExecuteAsync(cn).Wait();

                    // region Name + Id combos should be the same 
                    var rawRegions = cn.Query<NameId>("SELECT [Id], [Region] AS [Name] FROM [SalesDataRaw] ORDER BY [Id]");
                    var keyedRegions = cn.Query<NameId>("SELECT [d].[Id], [r].[Name] FROM [SalesDataKeyed] [d] INNER JOIN [Region] [r] ON [r].[Id]=[d].[RegionId] ORDER BY [d].[Id]");
                    Assert.IsTrue(rawRegions.SequenceEqual(keyedRegions));

                    // but there is deliberately one error where Type = Unknown
                    var errors = inlineLookup.GetErrorsAsync(cn).Result;
                    Assert.IsTrue(errors.Count == 1);
                    Assert.IsTrue(errors["Type"].SequenceEqual(new string[] { "Unknown" }));
                }
            }

            void CreateRows(SqlConnection cn, string tableName, params string[] names)
            {
                names.ToList().ForEach(name => cn.Execute($"INSERT INTO {tableName} ([Name]) VALUES (@name)", new { name }));
            }
        }

        [TestMethod]
        public async Task BulkCopy()
        {
            using (var cn = LocalDb.GetConnection(dbName))
            {
                await cn.DropAllTablesAsync();

                using (var stream = GetResource("zipcode_us_geo_db.xlsx"))
                {
                    var loader = new ExcelLoader();
                    var result = await loader.BulkSaveAsync(stream, cn, "dbo", "ZipCodes");

                    Assert.IsTrue(result > 80_237);
                }
            }
        }

        private IEnumerable<InitializeStatement> SampleLookupObjects()
        {
            yield return new InitializeStatement("dbo.Region", "DROP TABLE %obj%", @"CREATE TABLE %obj% (
                [Id] int identity(1,1) PRIMARY KEY,
                [Name] nvarchar(50) NOT NULL
            )");

            yield return new InitializeStatement("dbo.Type", "DROP TABLE %obj%", @"CREATE TABLE %obj% (
                [Id] int identity(1,1) PRIMARY KEY,
                [Name] nvarchar(50) NOT NULL
            )");
        }

        private static void ExecuteIgnoreError(IDbConnection connection, string command)
        {
            try
            {
                connection.Execute(command);
            }
            catch
            {
                // do nothing
            }
        }

        private static void CreateSchema(SqlConnection cn)
        {
            try { cn.Execute("CREATE SCHEMA [upload]"); } catch { /* do nothing */ }
        }

        private Stream GetResource(string resourceName)
        {
            //var names = Assembly.GetExecutingAssembly().GetManifestResourceNames();

            return Assembly.GetExecutingAssembly().GetManifestResourceStream($"Testing.Resources.{resourceName}");
        }

        private class NameId : IEquatable<NameId>
        {
            public int Id { get; set; }
            public string Name { get; set; }

            public bool Equals([AllowNull] NameId other) => Id.Equals(other.Id) && Name.Equals(other.Name);            
        }
    }
}
