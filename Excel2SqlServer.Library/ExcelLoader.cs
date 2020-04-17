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

namespace Excel2SqlServer.Library
{
    public class ExcelLoader
	{
		public void CreateTable(string fileName, SqlConnection connection, string schemaName, string tableName, IEnumerable<string> customColumns = null)
		{
			var ds = Read(fileName);
			CreateTableInner(ds, connection, schemaName, tableName, customColumns);
		}

		public void CreateTable(Stream stream, SqlConnection connection, string schemaName, string tableName, IEnumerable<string> customColumns = null)
		{
			var ds = Read(stream);
			CreateTableInner(ds, connection, schemaName, tableName, customColumns);
		}

		private void CreateTableInner(DataSet ds, SqlConnection connection, string schemaName, string tableName, IEnumerable<string> customColumns)
		{
			var loadTable = ds.Tables[0];

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

			IEnumerable<string> getColumns(DataTable dataTable)
			{
				yield return "[Id] int identity(1,1) PRIMARY KEY";

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

			string createCmd = $"CREATE TABLE [{schemaName}].[{tableName}] (\r\n{string.Join(",\r\n", getColumns(loadTable))}\r\n)";

			using (var cmd = new SqlCommand(createCmd, connection))
			{
				if (connection.State == ConnectionState.Closed) connection.Open();
				cmd.ExecuteNonQuery();
			}
		}

		public int Save(string fileName, SqlConnection connection, string schemaName, string tableName, Options options = null)
		{
			var ds = Read(fileName);
			return SaveInner(connection, schemaName, tableName, ds, options);
		}

		public int Save(Stream stream, SqlConnection connection, string schemaName, string tableName, Options options = null)
		{
			var ds = Read(stream);
			return SaveInner(connection, schemaName, tableName, ds, options);
		}

		private int SaveInner(SqlConnection connection, string schemaName, string tableName, DataSet ds, Options options)
        {
			if (!connection.TableExists(schemaName, tableName)) CreateTableInner(ds, connection, schemaName, tableName, options?.CustomColumns);
			SaveDataTable(connection, ds.Tables[0], schemaName, tableName, options);
			return ds.Tables[0].Rows.Count;
		}

		private void SaveDataTable(SqlConnection connection, DataTable table, string schemaName, string tableName, Options options)
		{
			if (options?.TruncateFirst ?? false) connection.Execute($"TRUNCATE TABLE [{schemaName}].[{tableName}]");

			// thanks to https://stackoverflow.com/a/4582786/2023653
			foreach (DataRow row in table.Rows)
			{
				row.AcceptChanges();
				row.SetAdded();
			}

			using (SqlCommand select = BuildSelectCommand(table, connection, schemaName, tableName))
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
				var columns = GetVarcharColumns(connection, schemaName, tableName);

				if (options?.AutoTrimStrings ?? false)
				{
					var updateCmds = BuildTrimCommands(columns, schemaName, tableName);
					foreach (var cmd in updateCmds) connection.Execute(cmd);
				}

				if (options?.RemoveNonPrintingChars ?? false)
				{
					const string nonPrintingChars = @"[^\x00-\x7F]+";
					var dataTable = connection.QueryTable($"SELECT * FROM [{schemaName}].[{tableName}]");
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
								$"UPDATE [{schemaName}].[{tableName}] SET {string.Join(", ", expressions.Select(kp => $"[{kp.Key}]='{kp.Value}'"))} WHERE [Id]=@id",
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

		public DataSet Read(string fileName)
		{
			using (var stream = File.OpenRead(fileName))
			{
				return Read(stream);
			}
		}

		public DataSet Read(Stream stream)
		{
			using (var reader = ExcelReaderFactory.CreateReader(stream))
			{
				return reader.AsDataSet(new ExcelDataSetConfiguration()
				{
					UseColumnDataType = true,
					ConfigureDataTable = (r) =>
					{
						return new ExcelDataTableConfiguration() { UseHeaderRow = true };
					}
				});
			}
		}
	}
}
