using Dapper;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace Testing.Extensions
{
    internal static class SqlConnectionExtensions
    {
        /// <summary>
        /// intended for use in integration test initialization to ensure you have a clean database
        /// </summary>
        public static async Task DropAllTablesAsync(this SqlConnection connection)
        {
            var foreignKeys = await connection.QueryAsync<AllForeignKeysResult>(
                @"SELECT 
	                SCHEMA_NAME([t].[schema_id]) AS [Schema],
	                [t].[name] AS [TableName],
	                [fk].[name] AS [ConstraintName]
                FROM 
	                [sys].[foreign_keys] [fk]
	                INNER JOIN [sys].[tables] [t] ON [fk].[parent_object_id]=[t].[object_id]");

            foreach (var fk in foreignKeys)
            {
                await connection.ExecuteAsync($"ALTER TABLE [{fk.Schema}].[{fk.TableName}] DROP CONSTRAINT [{fk.ConstraintName}]");
            }

            var tables = await connection.QueryAsync<AllTablesResult>(
                @"SELECT 
                    SCHEMA_NAME([t].[schema_id]) AS [Schema],
                    [t].[name] AS [TableName]
                FROM     
                    [sys].[tables] [t]");

            foreach (var t in tables)
            {
                await connection.ExecuteAsync($"DROP TABLE [{t.Schema}].[{t.TableName}]");
            }
        }

        private class AllForeignKeysResult
        {
            public string Schema { get; set; }
            public string TableName { get; set; }
            public string ConstraintName { get; set; }
        }

        private class AllTablesResult
        {
            public string Schema { get; set; }
            public string TableName { get; set; }
        }

    }
}
