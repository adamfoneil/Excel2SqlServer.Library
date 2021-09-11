using AO.Models;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace Excel2SqlServer.Library
{
    public partial class ExcelLoader
    {
        public async Task<int> BulkSaveAsync(Stream stream, SqlConnection connection, string schemaName, string tableName)
        {
            var ds = await ReadAsync(stream);
            return await BulkSaveInnerAsync(ds, connection, schemaName, tableName);
        }

        public async Task<int> BulkSaveAsync(string fileName, SqlConnection connection, string schemaName, string tableName)
        {
            var ds = await ReadAsync(fileName);
            return await BulkSaveInnerAsync(ds, connection, schemaName, tableName);
        }

        private async Task<int> BulkSaveInnerAsync(DataSet ds, SqlConnection connection, string schemaName, string tableName)
        {
            CreateTablesInner(ds, connection, new Dictionary<string, ObjectName>()
            {
                [ds.Tables[0].TableName] = new ObjectName(schemaName, tableName)
            }, null);

            using (var bcp = new SqlBulkCopy(connection))
            {
                var table = ds.Tables[0];
                foreach (DataColumn col in table.Columns) bcp.ColumnMappings.Add(col.ColumnName, col.ColumnName);

                bcp.DestinationTableName = $"{schemaName}.{tableName}";

                await bcp.WriteToServerAsync(table);

                return bcp.RowsCopied;
            }
        }
    }
}
