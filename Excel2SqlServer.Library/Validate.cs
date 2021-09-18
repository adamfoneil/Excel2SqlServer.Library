using AO.Models;
using Dapper;
using Excel2SqlServer.Library.Extensions;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace Excel2SqlServer.Library
{
    public static class Validate
    {
        /// <summary>
        /// Loops through records in a query and attempts conversions to specified types for the specified columns,
        /// and returns info about conversion failures
        /// </summary>
        public static IEnumerable<ValidationInfo> ColumnTypes(
            SqlConnection connection, string query, string reportColumn, params TypeValidator[] columns)
        {
            using (var cmd = new SqlCommand(query, connection))
            {
                using (var adapter = new SqlDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    return ColumnTypes(table, reportColumn, columns);
                }
            }
        }

        public static IEnumerable<ValidationInfo> ColumnTypes(
            DataTable dataTable, string reportColumn, params TypeValidator[] columns)
        {
            List<ValidationInfo> results = new List<ValidationInfo>();

            foreach (DataRow row in dataTable.Rows)
            {
                foreach (var col in columns)
                {
                    if (!col.TryConversion(row, out ValidationInfo info))
                    {
                        info.ReportValue = row[reportColumn].ToString();
                        results.Add(info);
                    }
                }
            }

            return results;
        }

        public static async Task<IEnumerable<ValidationInfo>> SqlServerTypeConversionAsync<TKey, TValue>(
            SqlConnection connection, string schema, string table, string keyColumn, string convertColumn, string convertType,
            string criteria = null)
        {
            string whereClause = (!string.IsNullOrEmpty(criteria)) ? " WHERE " + criteria : string.Empty;
            var keys = await connection.QueryAsync<TKey>($"SELECT [{keyColumn}] FROM [{schema}].[{table}]{whereClause}");

            List<ValidationInfo> results = new List<ValidationInfo>();

            string testQuery = $"SELECT CONVERT({convertType}, [{convertColumn}]) AS [ConvertedValue] FROM [{schema}].[{table}] WHERE [{keyColumn}]=@key";
            string srcValueQuery = $"SELECT [{convertColumn}] AS [SourceValue] FROM [{schema}].[{table}] WHERE [{keyColumn}]=@key";

            foreach (var key in keys)
            {
                try
                {
                    var test = await connection.QuerySingleAsync<ConversionTest>(testQuery, new { key });
                }
                catch (Exception exc)
                {
                    object offendingValue = null;

                    try
                    {
                        offendingValue = await connection.QuerySingleOrDefaultAsync<TValue>(srcValueQuery, new { key });
                    }
                    catch (Exception excInner)
                    {
                        offendingValue = "Couldn't determine offending value: " + excInner.Message;
                    }

                    results.Add(new ValidationInfo()
                    {
                        ReportValue = key.ToString(),
                        ColumnName = convertColumn,
                        Message = exc.Message,
                        OffendingValue = offendingValue
                    });
                }
            }

            return results;
        }

        /// <summary>
        /// gets columns in a source table with data that's too long for mapped destination columns.
        /// columnMappings argument is source column to destination
        /// </summary>
        public static async Task<Dictionary<string, (string value, int data, int allowed)>> GetOversizedDataAsync(
            SqlConnection connection, ObjectName source, ObjectName destination, Dictionary<string, string> columnMappings)
        {                                    
            Dictionary<string, (string, int)> maxLengths = await GetMaxDataLengthsAsync(connection, source);

            return await GetOversizedDataAsync(connection, maxLengths, destination, columnMappings);
        }
        
        public static async Task<Dictionary<string, (string value, int length, int allowed)>> GetOversizedDataAsync(
            SqlConnection connection, Dictionary<string, (string value, int length)> maxDataLengths, ObjectName destination, Dictionary<string, string> columnMappings)
        {
            Dictionary<string, int> columnSizes = await GetColumnSizesAsync(connection, destination);

            return maxDataLengths
                .Where(kp => 
                {
                    if (!columnMappings.ContainsKey(kp.Key)) return false;
                    var sourceCol = columnMappings[kp.Key];
                    return kp.Value.length > columnSizes[sourceCol];
                })
                .ToDictionary(kp => kp.Key, kp => (kp.Value.value, kp.Value.length, columnSizes[columnMappings[kp.Key]]));
        }

        public static async Task EnsureNoOversizedDataAsync(
            SqlConnection connection, ObjectName source, ObjectName destination, Dictionary<string, string> columnMappings)
        {
            Dictionary<string, (string value, int length)> maxLengths = await GetMaxDataLengthsAsync(connection, source);

            await EnsureNoOversizedDataAsync(connection, maxLengths, destination, columnMappings);
        }

        public static async Task EnsureNoOversizedDataAsync(
            SqlConnection connection, Dictionary<string, (string value, int length)> maxDataLengths, ObjectName destination, Dictionary<string, string> columnMappings)
        {
            var oversized = await GetOversizedDataAsync(connection, maxDataLengths, destination, columnMappings);

            var sourceMappings = columnMappings.ToDictionary(kp => kp.Value, kp => kp.Key);

            if (oversized.Any())
            {
                var message = string.Join("\r\n", oversized.Select(kp => $"Data value '{kp.Value.value}' with length {kp.Value.length} in {sourceMappings[columnMappings[kp.Key]]} column can't insert into {destination.Schema}.{destination.Name}.{columnMappings[kp.Key]} due to max length {kp.Value.allowed}"));
                var exc = new Exception(message);
                foreach (var kp in oversized) exc.Data.Add(kp.Key, kp.Value);
                throw exc;
            }
        }

        private static async Task<Dictionary<string, int>> GetColumnSizesAsync(SqlConnection connection, ObjectName table)
        {
            var data = await connection.QueryAsync<ColumnSizeInfoResult>(
                @"DECLARE @sizers TABLE (
                    [Name] nvarchar(50) NOT NULL,
                    [Divisor] int NOT NULL
                )

                INSERT INTO @sizers ([Name], [Divisor]) 
                VALUES ('varchar', 1), ('nvarchar', 2), ('varbinary', 1), ('char', 1), ('nchar', 2)

                SELECT
                    [col].[name] AS [ColumnName],
                    [col].[max_length] / [s].[Divisor] AS [MaxSize]
                FROM
                    [sys].[columns] [col]
                    INNER JOIN @sizers [s] ON TYPE_NAME([col].[system_type_id])=[s].[Name]
                WHERE
                    [object_id]=OBJECT_ID(@objectName) AND    
                    [max_length] > -1", new
                {
                    objectName = table.ToString()
                });

            return data.ToDictionary(row => row.ColumnName, row => row.MaxSize);
        }

        public static async Task<Dictionary<string, (string value, int length)>> GetMaxDataLengthsAsync(SqlConnection connection, ObjectName table)
        {
            var columnNames = await connection.GetColumnNamesAsync(table);

            var results = new Dictionary<string, (string, int)>();

            foreach (var col in columnNames)
            {
                var dataLength = await connection.QuerySingleOrDefaultAsync<MaxLengthData>(
                    $@"WITH [max_len] AS (
                        SELECT MAX(LEN([{col}])) AS [Length] FROM [{table.Schema}].[{table.Name}]
                    ) SELECT TOP (1)
                        [src].[{col}] AS [Value], [ml].[Length]
                    FROM
                        [{table.Schema}].[{table.Name}] [src]
                        INNER JOIN [max_len] [ml] ON [ml].[Length]=LEN([src].[{col}])");

                if (dataLength != null)
                {
                    results.Add(col, (dataLength.Value, dataLength.Length));
                }                
            }

            return results;
        }

        private class MaxLengthData
        {
            public string Value { get; set; }
            public int Length { get; set; }
        }

        public class ColumnSizeInfoResult
        {
            public string ColumnName { get; set; }
            public int MaxSize { get; set; }
        }
    }

    internal class ConversionTest
    {
        public object ConvertedValue { get; set; }
        public object SourceValue { get; set; }
    }

    public class TypeValidator
    {
        public TypeValidator(string columnName, Type type)
        {
            ColumnName = columnName;
            Type = type;
        }

        public string ColumnName { get; }
        public Type Type { get; }

        public bool TryConversion(DataRow dataRow, out ValidationInfo info)
        {
            // if this fails, then you have a bad column name,
            // which is outside the scope of what we're trying to validate
            object result = dataRow[ColumnName];

            try
            {
                var typedResult = Convert.ChangeType(result, Type);
                info = null;
                return true;
            }
            catch (Exception exc)
            {
                info = new ValidationInfo()
                {
                    ColumnName = ColumnName,
                    Message = $"{Type.Name} conversion on column {ColumnName} failed: {exc.Message}",
                    OffendingValue = result
                };
                return false;
            }
        }
    }

    public class ValidationInfo
    {
        /// <summary>
        /// Reference value to help you find data in whatever your source is (must be convertable to string)
        /// </summary>
        public string ReportValue { get; set; }

        /// <summary>
        /// Column with the offending value
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Source value that couldn't be converted
        /// </summary>
        public object OffendingValue { get; set; }

        /// <summary>       
        /// Conversion error message
        /// </summary>
        public string Message { get; set; }
    }
}
