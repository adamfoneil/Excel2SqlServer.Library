using Dapper;
using DataTables.Library;
using Excel2SqlServer.Library.Extensions;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel2SqlServer.Library
{
    public class ExcelLoader
	{
		public async Task CreateTableAsync(string fileName, SqlConnection connection, string schemaName, string tableName, IEnumerable<string> customColumns = null)
        {
            var ds = await ReadAsync(fileName);
            CreateTablesInner(ds, connection, GetObjNameDictionary(ds, schemaName, tableName), customColumns);
        }

        public async Task CreateTableAsync(Stream stream, SqlConnection connection, string schemaName, string tableName, IEnumerable<string> customColumns = null)
		{
			var ds = await ReadAsync(stream);
			CreateTablesInner(ds, connection, GetObjNameDictionary(ds, schemaName, tableName), customColumns);
		}

		/// <summary>
		/// for backward compatibility with methods that assumed single worksheet/single table load
		/// </summary>
		private static Dictionary<string, ObjectName> GetObjNameDictionary(DataSet ds, string schemaName, string tableName)
		{
			return new Dictionary<string, ObjectName>()
			{
				{  ds.Tables[0].TableName, new ObjectName() { Schema = schemaName, Name = tableName } }
			};
		}

		private static Dictionary<string, ObjectName> GetDefaultTableNaming(DataSet ds, string schema)
        {
			return ds.Tables.OfType<DataTable>().ToDictionary(tbl => tbl.TableName, tbl => new ObjectName() { Schema = schema, Name = tbl.TableName });
        }

		private void CreateTablesInner(DataSet ds, SqlConnection connection, Dictionary<string, ObjectName> tableNames, IEnumerable<string> customColumns)
		{
			Dictionary<Type, string> typeMappings = new Dictionary<Type, string>()
			{
				{ typeof(object), "nvarchar(max)" },
				{ typeof(int), "int" },
				{ typeof(string), "nvarchar(max)" },
				{ typeof(DateTime), "datetime" },
				{ typeof(bool), "bit" },
				{ typeof(double), "float" },
				{ typeof(decimal), "decimal" }
			};

			int nameIndex = 0;
			foreach (DataTable tbl in ds.Tables)
            {
				nameIndex++;
				ObjectName objName = (tableNames.ContainsKey(tbl.TableName)) ?
					tableNames[tbl.TableName] :
					new ObjectName() { Schema = "dbo", Name = tbl.TableName + nameIndex.ToString() };

				if (!connection.SchemaExists(objName.Schema))
                {
					execute($"CREATE SCHEMA [{objName.Schema}]");					
                }

				if (!connection.TableExists(objName.Schema, objName.Name))
                {
					execute($"CREATE TABLE [{objName.Schema}].[{objName.Name}] (\r\n{string.Join(",\r\n", getColumns(tbl))}\r\n)");
				}
			}

			void execute(string command)
            {
				using (var cmd = new SqlCommand(command, connection))
				{
					if (connection.State == ConnectionState.Closed) connection.Open();
					cmd.ExecuteNonQuery();
				}
			}

			IEnumerable<string> getColumns(DataTable dataTable)
			{
				const string identityCol = "Id";
				if (!dataTable.Columns.Contains(identityCol))
                {
					yield return $"[{identityCol}] int identity(1,1) PRIMARY KEY";
				}
				
				if (customColumns?.Any() ?? false)
				{
					foreach (string col in customColumns) yield return col;
				}

				foreach (DataColumn col in dataTable.Columns)
				{
					if (!typeMappings.ContainsKey(col.DataType)) throw new Exception($"ExcelLoader.CreateTableInner is missing a type mapping for data type {col.DataType.Name}");
					yield return $"[{col.ColumnName}] {typeMappings[col.DataType]} NULL";
				}
			};
		}

		public async Task<int> SaveAsync(string fileName, SqlConnection connection, Dictionary<string, ObjectName> tableNames = null, Options options = null)
        {
			var ds = await ReadAsync(fileName);
			return await SaveInnerAsync(connection, ds, tableNames ?? GetDefaultTableNaming(ds, options?.SchemaName ?? "dbo"), options);
		}

		public async Task<int> SaveAsync(string fileName, SqlConnection connection, string schemaName, string tableName, Options options = null)
		{
			var ds = await ReadAsync(fileName);
			return await SaveInnerAsync(connection, ds, GetObjNameDictionary(ds, schemaName, tableName), options);
		}

		public async Task<int> SaveAsync(Stream stream, SqlConnection connection, Dictionary<string, ObjectName> tableNames = null, Options options = null)
        {
			var ds = await ReadAsync(stream);
			return await SaveInnerAsync(connection, ds, tableNames ?? GetDefaultTableNaming(ds, options?.SchemaName ?? "dbo"), options);
		}

		public async Task<int> SaveAsync(Stream stream, SqlConnection connection, string schemaName, string tableName, Options options = null)
		{
			var ds = await ReadAsync(stream);
			return await SaveInnerAsync(connection, ds, GetObjNameDictionary(ds, schemaName, tableName), options);
		}

		private async Task<int> SaveInnerAsync(SqlConnection connection, DataSet ds, Dictionary<string, ObjectName> tableNames, Options options)
        {
			CreateTablesInner(ds, connection, tableNames, options?.CustomColumns);

			int count = 0;
			await Task.Run(() =>
			{
				foreach (var tableName in tableNames)
                {
					DataTable tbl = ds.Tables[tableName.Key];
					SaveDataTable(connection, tbl, tableNames[tbl.TableName], options);
					count += tbl.Rows.Count;
				}								
			});
			return count;
		}

		private void SaveDataTable(SqlConnection connection, DataTable table, ObjectName objName, Options options)
		{
			if (options?.TruncateFirst ?? false) connection.Execute($"TRUNCATE TABLE [{objName.Schema}].[{objName.Name}]");

			// thanks to https://stackoverflow.com/a/4582786/2023653
			foreach (DataRow row in table.Rows)
			{
				row.AcceptChanges();
				row.SetAdded();
			}

			using (SqlCommand select = BuildSelectCommand(table, connection, objName.Schema, objName.Name))
			{
				using (var adapter = new SqlDataAdapter(select))
				{
					using (var builder = new SqlCommandBuilder(adapter))
					{
						adapter.InsertCommand = builder.GetInsertCommand();
						adapter.Update(table);
					}
				}
			}

			if ((options?.AutoTrimStrings ?? false) || (options?.RemoveNonPrintingChars ?? false))
			{
				var columns = GetVarcharColumns(connection, objName.Schema, objName.Name);

				if (options?.AutoTrimStrings ?? false)
				{
					var updateCmds = BuildTrimCommands(columns, objName.Schema, objName.Name);
					foreach (var cmd in updateCmds) connection.Execute(cmd);
				}

				if (options?.RemoveNonPrintingChars ?? false)
				{
					const string nonPrintingChars = @"[^\x00-\x7F]+";
					var dataTable = connection.QueryTable($"SELECT * FROM [{objName.Schema}].[{objName.Name}]");
					foreach (DataRow row in dataTable.Rows)
					{
						var expressions = new Dictionary<string, string>();						

						foreach (var col in columns)
						{
							if (!row.IsNull(col) && Regex.IsMatch(row.Field<string>(col), nonPrintingChars))
							{
								expressions.Add(col, Regex.Replace(row.Field<string>(col), nonPrintingChars, string.Empty).Replace("'", "''"));
							}
						}

						if (expressions.Any())
						{
							connection.Execute(
								$"UPDATE [{objName.Schema}].[{objName.Name}] SET {string.Join(", ", expressions.Select(kp => $"[{kp.Key}]='{kp.Value}'"))} WHERE [Id]=@id",
								new { id = row.Field<int>("Id") });
						}
					}
				}
			}
		}

		/// <summary>
		/// returns update statements for each varchar column in specified table that calls LTRIM(RTRIM()) on varchar columns
		/// </summary>
		private IEnumerable<string> BuildTrimCommands(IEnumerable<string> columns, string schemaName, string tableName)
		{
			foreach (var col in columns)
			{
				yield return $"UPDATE [{schemaName}].[{tableName}] SET [{col}]=LTRIM(RTRIM([{col}])) WHERE [{col}] IS NOT NULL";
			}
		}

		private static IEnumerable<string> GetVarcharColumns(SqlConnection connection, string schemaName, string tableName)
		{
			return connection.Query<string>(
				@"SELECT 
					[col].[name]
				FROM 
					[sys].[columns] [col] INNER JOIN [sys].[tables] [t] ON [col].[object_id]=[t].[object_id]
				WHERE 
					SCHEMA_NAME([t].[schema_id])=@schemaName AND 
					[t].[name]=@tableName AND
					TYPE_NAME([col].[system_type_id]) IN ('varchar', 'nvarchar')", new { schemaName, tableName });
		}

		private SqlCommand BuildSelectCommand(DataTable table, SqlConnection connection, string schemaName, string tableName)
		{
			string[] columnNames = table.Columns.OfType<DataColumn>().Select(col => col.ColumnName).ToArray();
			string query = $"SELECT {string.Join(", ", columnNames.Select(col => $"[{col}]"))} FROM [{schemaName}].[{tableName}]";
			return new SqlCommand(query, connection);
		}

		public async Task<DataSet> ReadAsync(string fileName)
		{
			using (var stream = File.OpenRead(fileName))
			{
				return await ReadAsync(stream);
			}
		}

		public async Task<DataSet> ReadAsync(Stream stream)
		{
			DataSet result = null;

			await Task.Run(() =>
			{
				using (var reader = ExcelReaderFactory.CreateReader(stream))
				{
					result = reader.AsDataSet(new ExcelDataSetConfiguration()
					{
						UseColumnDataType = true,
						ConfigureDataTable = (r) =>
						{
							return new ExcelDataTableConfiguration() { UseHeaderRow = true };
						}
					});
				}
			});

			return result;
		}

		public class ObjectName
        {
            public ObjectName()
            {
            }

            public ObjectName(string schema, string name)
            {
				Schema = schema;
				Name = name;
            }

			public string Schema { get; set; }
			public string Name { get; set; }
        }
	}
}
