using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace Excel2SqlServer.Library
{
    public static class Validation
    {
        /// <summary>
        /// Loops through records in a query and attempts conversions to specified types for the specified columns,
        /// and returns info about conversion failures
        /// </summary>
        public static IEnumerable<ValidationInfo> ValidateColumnTypes(
            SqlConnection connection, string query, string reportColumn, params TypeValidator[] columns)
        {
            using (var cmd = new SqlCommand(query, connection))
            {
                using (var adapter = new SqlDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    return ValidateColumnTypes(table, reportColumn, columns);
                }
            }
        }

        public static IEnumerable<ValidationInfo> ValidateColumnTypes(
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

        public static async Task<IEnumerable<ValidationInfo>> ValidateSqlServerTypeConversionAsync<TKey, TValue>(
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
