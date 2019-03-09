using Dapper;
using Excel2SqlServer.Library.Extensions;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

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
				{ typeof(bool), "bit" }
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

		public int Save(string fileName, SqlConnection connection, string schemaName, string tableName, bool truncateFirst = false, IEnumerable<string> customColumns = null)
		{
			var ds = Read(fileName);
			return SaveInner(connection, schemaName, tableName, truncateFirst, customColumns, ds);
		}

		public int Save(Stream stream, SqlConnection connection, string schemaName, string tableName, bool truncateFirst = false, IEnumerable<string> customColumns = null)
		{
			var ds = Read(stream);
			return SaveInner(connection, schemaName, tableName, truncateFirst, customColumns, ds);
		}

		private int SaveInner(SqlConnection connection, string schemaName, string tableName, bool truncateFirst, IEnumerable<string> customColumns, DataSet ds)
		{
			if (!connection.TableExists(schemaName, tableName)) CreateTableInner(ds, connection, schemaName, tableName, customColumns);
			SaveDataTable(connection, ds.Tables[0], schemaName, tableName, truncateFirst);
			return ds.Tables[0].Rows.Count;
		}

		private void SaveDataTable(SqlConnection connection, DataTable table, string schemaName, string tableName, bool truncateFirst)
		{
			if (truncateFirst) connection.Execute($"TRUNCATE TABLE [{schemaName}].[{tableName}]");

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
