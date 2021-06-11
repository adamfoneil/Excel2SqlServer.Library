using AO.Models;
using Dapper;
using DataTables.Library;
using Excel2SqlServer.Library.Extensions;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace Excel2SqlServer.Library
{
    public class InlineLookup<T>
    {        
        public InlineLookup(string sourceTable, string identityColumn, string resultTable, IEnumerable<Lookup> lookups)
        {
            SourceTable = sourceTable;
            IdentityColumn = identityColumn;
            ResultTable = resultTable;
            Lookups = lookups.ToDictionary(item => item.SourceColumn);
        }

        public string SourceTable { get; }
        public string IdentityColumn { get; }
        public string ResultTable { get; }
        public Dictionary<string, Lookup> Lookups { get; }

        /// <summary>
        /// builds the result table. NULLs in the output mean no match was found
        /// </summary>
        public async Task ExecuteAsync(SqlConnection connection)
        {
            await RebuildResultTableAsync(connection);

            // create blank rows in the result table for each source row.
            // this is what we'll update with key values next
            var sourceObj = ObjectName.FromName(SourceTable);
            var resultObj = ObjectName.FromName(ResultTable);
            var columnNames = string.Join(", ", Lookups.Select(kp => $"[{kp.Value.ResultColumn}]"));
            var defaultValues = string.Join(", ", Lookups.Select(kp => "NULL"));
            await connection.ExecuteAsync(
                $@"INSERT INTO [{resultObj.Schema}].[{resultObj.Name}] ([{IdentityColumn}], {columnNames})
                SELECT [{IdentityColumn}], {defaultValues}
                FROM [{sourceObj.Schema}].[{sourceObj.Name}]");
            
            // attempt to match text values in the source table with named values in the corresponding lookup table.
            // if no match, we simply get nulls in the corresponding output col
            foreach (var col in Lookups)
            {
                var lookupObj = ObjectName.FromName(col.Value.LookupTable);
                var sqlUpdate =
                    $@"UPDATE [result] SET 
                        [{col.Value.ResultColumn}]=[lookup].[{col.Value.LookupIdentityColumn}]
                    FROM 
                        [{sourceObj.Schema}].[{sourceObj.Name}] [src]
                        INNER JOIN [{resultObj.Schema}].[{resultObj.Name}] [result] ON [src].[{IdentityColumn}]=[result].[{IdentityColumn}]
                        INNER JOIN [{lookupObj.Schema}].[{lookupObj.Name}] [lookup] ON [src].[{col.Key}]=[lookup].[{col.Value.LookupNameColumn}]";

                await connection.ExecuteAsync(sqlUpdate);
            }
        }

        /// <summary>
        /// what's the SQL syntax for type T?       
        /// </summary>        
        protected virtual string ColumnSqlTypeSyntax =>
            (typeof(T).Equals(typeof(int))) ? "int" : 
            throw new Exception($"Type {typeof(T).Name} is not currently supported by InlineLookup.");

        private async Task RebuildResultTableAsync(SqlConnection connection)
        {
            var objName = ObjectName.FromName(ResultTable);
            if (connection.TableExists(objName.Schema, objName.Name))
            {
                await connection.ExecuteAsync($"DROP TABLE [{objName.Schema}].[{objName.Name}]");
            }

            string sql = $"CREATE TABLE [{objName.Schema}].[{objName.Name}] (";

            List<string> columns = new List<string>();

            // this is what joins to the source table (spreadsheet uploaded by user with text key values we're going to lookup)
            // it's hardcoded int because ExcelLoader.cs:93
            columns.Add($"[{IdentityColumn}] int NOT NULL PRIMARY KEY");

            // now add the lookup columns -- this is where the key lookup results are stored
            columns.AddRange(Lookups.Select(kp => $"[{kp.Value.ResultColumn}] {ColumnSqlTypeSyntax} NULL"));

            sql += string.Join("\r\n\t, ", columns) + "\r\n)";

            await connection.ExecuteAsync(sql);
        }

        /// <summary>
        /// finds source values that don't have a mapping
        /// </summary>
        public async Task<ILookup<string, string>> GetErrorsAsync(SqlConnection cn)
        {
            throw new NotImplementedException();
        }
    }

    public class Lookup
    {
        public Lookup()
        {
        }

        public Lookup(string sourceColumn, string resultColumn, string lookupTable, string lookupNameColumn, string lookupIdentityColumn)
        {
            SourceColumn = sourceColumn;
            ResultColumn = resultColumn;
            LookupTable = lookupTable;
            LookupNameColumn = lookupNameColumn;
            LookupIdentityColumn = lookupIdentityColumn;
        }

        /// <summary>
        /// string column in the source table being converted to T
        /// For example RegionName
        /// </summary>
        public string SourceColumn { get; set; }

        /// <summary>
        /// T column being converted into (in the ResultTable), typically a numeric key value column
        /// For example RegionId
        /// </summary>
        public string ResultColumn { get; set; }

        /// <summary>
        /// what table are we joining to to get a key value
        /// </summary>
        public string LookupTable { get; set; }

        /// <summary>
        /// What column do we join to in the lookup table to get the identity value
        /// </summary>
        public string LookupNameColumn { get; set; }

        /// <summary>
        /// what column are we returning from the lookup table?
        /// This is what we're ultimately trying to find
        /// </summary>
        public string LookupIdentityColumn { get; set; }
    }
}
