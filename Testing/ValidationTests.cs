using Excel2SqlServer.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Postulate.SqlServer;
using SqlServer.LocalDb;
using SqlServer.LocalDb.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace Testing
{
    [TestClass]
    public class ValidationTests
    {
        private const string DateValidatonTable = "dbo.DateValidation";

        [TestMethod]
        public void ValidateSqlServerTypeConversionAsyncDateTime()
        {
            using (var cn = LocalDb.GetConnection("sample", CreateObjects()))
            {
                // create a mix of valid and invalid date values as strings

                var validDates = new string[]
                {
                    "3/6/19",
                    "4/12/16",
                    "1/19/08",
                    "9/29/12",
                    "10/17/14",
                    "5/18/21",
                    "8/25/18",
                    "10/02/11",
                    "2/29/04"
                };

                var invalidDates = new string[]
                {
                    "2/38/07",
                    "7/12/290",
                    "6/31/09",
                    "2/29/03",
                    "14/1/05",
                    "3/4/1634" // note this is a valid date, but not a valid datetime. I'm checking datetime in this test
                };

                // dates are loaded as strings, so they are all "valid" at this point
                CreateDateValues(cn, validDates.Concat(invalidDates));

                // now we find the actual invalid datetimes
                var results = Validation.ValidateSqlServerTypeConversionAsync<string, string>(cn, "dbo", "DateValidation", "ProposedDate", "ProposedDate", "datetime").Result;

                // the "offending values" should be exactly the same as the invalidDates above
                var discoveredValuesSorted = results.Select(info => info.OffendingValue.ToString()).OrderBy(s => s).ToArray();
                var sourceValuesSorted = invalidDates.OrderBy(s => s).ToArray();

                // need to compare arrays in the same order, hence the sorting
                Assert.IsTrue(discoveredValuesSorted.SequenceEqual(sourceValuesSorted));
            }
        }

        private void CreateDateValues(SqlConnection cn, IEnumerable<string> dateStrings)
        {
            foreach (var value in dateStrings)
            {
                new SqlServerCmd(DateValidatonTable, "Id")
                {
                    { "ProposedDate", value }
                }.InsertAsync<int>(cn).Wait();
            }
        }

        private IEnumerable<InitializeStatement> CreateObjects()
        {
            yield return new InitializeStatement(
                DateValidatonTable,
                "DROP TABLE %obj%",
                @"CREATE TABLE %obj% (
                    [ProposedDate] varchar(20) NOT NULL PRIMARY KEY,
                    [Id] int identity(1,1)
                )");
        }
    }
}
