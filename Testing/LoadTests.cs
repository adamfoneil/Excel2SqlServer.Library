using Dapper;
using Excel2SqlServer.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SqlServer.LocalDb;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;

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
                    new ExcelLoader().SaveAsync(stream, cn, new Dictionary<string, ExcelLoader.ObjectName>()
                    {
                        { "Companies", new ExcelLoader.ObjectName("loader", "Company") },
                        { "People", new ExcelLoader.ObjectName("loader", "Person") }
                    }).Wait();

                    Assert.IsTrue(cn.Query<int>("SELECT COUNT(1) FROM [loader].[Company]").Count() > 0);
                    Assert.IsTrue(cn.Query<int>("SELECT COUNT(1) FROM [loader].[Person]").Count() > 0);
                }
            }
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
    }
}
