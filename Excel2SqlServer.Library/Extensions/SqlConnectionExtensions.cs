using AO.Models;
using Dapper;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Excel2SqlServer.Library.Extensions
{
    public static class SqlConnectionExtensions
    {
        public static bool TableExists(this SqlConnection connection, string schemaName, string tableName)
        {
            return (connection.QuerySingleOrDefault<int?>(
                "SELECT 1 FROM [sys].[tables] WHERE [schema_id]=SCHEMA_ID(@schemaName) AND [name]=@tableName",
                new { schemaName, tableName }) ?? 0) == 1;
        }

        public static bool SchemaExists(this SqlConnection connection, string schemaName)
        {
            return (connection.QuerySingleOrDefault<int?>(
                "SELECT 1 FROM [sys].[schemas] WHERE [name]=@schemaName", new { schemaName }) ?? 0) == 1;
        }

        public static async Task<IEnumerable<string>> GetColumnNamesAsync(this SqlConnection connection, string schemaName, string tableName) =>
            await GetColumnNamesAsync(connection, new ObjectName(schemaName, tableName));

        public static async Task<IEnumerable<string>> GetColumnNamesAsync(this SqlConnection connection, ObjectName table) =>
            await connection.QueryAsync<string>(
                @"SELECT [col].[name] FROM [sys].[columns] [col] WHERE [object_id]=OBJECT_ID(@objectName)",
                new { objectName = table.ToString() });
    }
}