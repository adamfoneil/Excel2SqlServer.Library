using Dapper;
using Excel2SqlServer.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SqlServer.LocalDb;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

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
                    loader.Save(stream, cn, "upload", "Case01");
                }
            }
        }

        [TestMethod]
        public void Case02()
        {
            using (var cn = LocalDb.GetConnection(dbName))
            {
                CreateSchema(cn);

                var loader = new ExcelLoader();
                using (var stream = GetResource("case02.xlsx"))
                {
                    loader.Save(stream, cn, "upload", "Case02", new Options()
                    {
                        TruncateFirst = true,
                        AutoTrimStrings = true,
                        RemoveNonPrintingChars = true
                    });
                }

                var lastNames = cn.Query<string>("SELECT [Last Name] FROM [upload].[Case02]");
                Assert.IsTrue(lastNames.All(name => name.Equals(name.Trim())));

                var addresses = cn.Query<string>("SELECT [Address] FROM [upload].[Case02]");
                Assert.IsTrue(addresses.All(address => address.Equals(address.Trim())));
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
