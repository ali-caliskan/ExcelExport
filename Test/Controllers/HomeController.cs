using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Test.Models;
using ClosedXML.Excel;
using System.Data;
using System.Configuration;
using System.Web.Configuration;

namespace Test.Controllers
{
	public class HomeController : Controller
	{
		private string connectionString = "data source=DESKTOP-SDMSMOQ\\SQLEXPRESS;initial catalog=Dbex;integrated security=True;";
		
		public ActionResult Index()
		{
			return View();
		}
		
		[HttpPost]
		public ActionResult UploadExcel(HttpPostedFileBase file)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			if (file != null && file.ContentLength > 0 && (file.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || file.ContentType == "application/vnd.ms-excel"))
			{
				using (var package = new ExcelPackage(file.InputStream))
				{
					var worksheet = package.Workbook.Worksheets[0];

					// Tablo oluşturma (Excel sütunlarına göre)
					string tableName = "DataTable";
					CreateTableFromExcel(worksheet, tableName);

					// Tablo varsa, tabloyu temizle
					ClearTable(tableName);

					// Verileri oku ve veritabanına ekle
					var data = ReadDataTable(worksheet);
					InsertDataTable(data, tableName);
				}
			}

			return RedirectToAction("Index");
			
		}

		private void ClearTable(string tableName)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				string query = $"DELETE FROM [{tableName}]";
				using (SqlCommand command = new SqlCommand(query, connection))
				{
					command.ExecuteNonQuery();
				}
			}
		}


		private void InsertDataTable(DataTable dataTable, string tableName)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();

				foreach (DataRow row in dataTable.Rows)
				{
					// Sütun adlarını DataTable'dan dinamik olarak al
					var columnNames = dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

					string insertQuery = $"INSERT INTO [{tableName}] ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", columnNames.Select(c => "@" + c))})";

					using (SqlCommand command = new SqlCommand(insertQuery, connection))
					{
						foreach (DataColumn column in dataTable.Columns)
						{
							var value = row[column];

							// Eğer değer null ise DBNull.Value kullan
							command.Parameters.AddWithValue("@" + column.ColumnName, value ?? DBNull.Value);
						}

						command.ExecuteNonQuery();
					}
				}
			}
		}

		private DataTable ReadDataTable(ExcelWorksheet worksheet)
		{
			DataTable dataTable = new DataTable();

			// İlk satırdan sütun adlarını al ve DataTable'a ekle
			for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
			{
				string columnName = worksheet.Cells[1, col].Value?.ToString().Replace(" ", "_");
				if (string.IsNullOrEmpty(columnName))
				{
					columnName = $"Column_{col}"; // Boş sütun adlarına otomatik isim verme
				}

				dataTable.Columns.Add(columnName, GetColumnDataType(worksheet.Cells[2, col].Value));
			}

			// Verileri DataTable'a ekle
			for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
			{
				DataRow dataRow = dataTable.NewRow();
				for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
				{
					var cellValue = worksheet.Cells[row, col].Value;
					dataRow[col - 1] = cellValue;
				}
				dataTable.Rows.Add(dataRow);
			}

			return dataTable;
		}

		private Type GetColumnDataType(object firstCellValue)
		{
			if (firstCellValue == null)
				return typeof(string); // Varsayılan olarak string

			Type cellType = firstCellValue.GetType();

			if (cellType == typeof(double))
				return typeof(double);
			else if (DateTime.TryParse(firstCellValue.ToString(), out _))
				return typeof(DateTime);
			else
				return typeof(string); // Diğer durumlarda string
		}

		private void CreateTableFromExcel(ExcelWorksheet worksheet, string tableName)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();

				if (TableExists(connection, tableName))
				{
					// Tablo varsa, ne yapılacağını belirle (örneğin, silme, güncelleme, temizleme veya farklı bir tablo adı kullanma)
				}
				else
				{
					StringBuilder createTableQuery = new StringBuilder($"CREATE TABLE [{tableName}] (");

					// İlk satırdaki değerleri sütun adları olarak kullan
					for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
					{
						string columnName = worksheet.Cells[1, col].Value?.ToString().Replace(" ", "_");
						if (string.IsNullOrEmpty(columnName))
						{
							columnName = $"Column_{col}";
						}
						string columnType = GetSqlDataType(worksheet.Cells[2, col].Value?.ToString() ?? "System.String");

						createTableQuery.Append($"[{columnName}] {columnType}, ");
					}

					createTableQuery.Length -= 2; // Son virgülü kaldır
					createTableQuery.Append(")");

					using (SqlCommand command = new SqlCommand(createTableQuery.ToString(), connection))
					{
						try
						{
							command.ExecuteNonQuery();
						}
						catch (SqlException ex)
						{
							throw new Exception($"Tablo oluşturma hatası: {ex.Message}");
						}
					}
				}
			}
		}

		private bool TableExists(SqlConnection connection, string tableName)
		{
			using (SqlCommand command = new SqlCommand($"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}'", connection))
			{
				int count = (int)command.ExecuteScalar();
				return count > 0;
			}
		}


		private string GetSqlDataType(string DataTableType)
		{
			switch (DataTableType)
			{
				case "System.Double":
					return "FLOAT";
				case "System.String":
					return "NVARCHAR(255)";
				case "System.DateTime":
					return "DATETIME";
				default:
					return "NVARCHAR(MAX)";
			}
		}



		public ActionResult About()
		{
			ViewBag.Message = "Your application description page.";

			return View();
		}

		public ActionResult Contact()
		{
			ViewBag.Message = "Your contact page.";

			return View();
		}
	}
}