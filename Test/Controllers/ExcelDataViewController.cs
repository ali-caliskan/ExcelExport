using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using System.IO;
using System.Configuration;

namespace Test.Controllers
{
	public class ExcelDataViewController : Controller
	{
		// GET: ExcelDataView
		private string connectionString = "data source=DESKTOP-SDMSMOQ\\SQLEXPRESS;initial catalog=Dbex;integrated security=True;";
		
		public ActionResult Index()
		{
			
			

			
				string query = "SELECT * FROM DataTable";


			DataTable dataTable = new DataTable();
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();

				using (SqlDataReader reader = command.ExecuteReader())
				{

					// Load the data from the reader into the DataTable.
					dataTable.Load(reader);
					
				}
			}

			// Pass the DataTable to the view.
			ViewBag.DataTable = dataTable;
			return View();
		}



		[HttpPost]
		public ActionResult ExportExcel()
		{
			DataTable dataTable = new DataTable();
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand("SELECT * FROM DataTable", connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();
				dataTable.Load(reader);
			}

			using (var workbook = new XLWorkbook())
			{
				var worksheet = workbook.Worksheets.Add("DataTable");
				worksheet.Cell(1, 1).InsertTable(dataTable);

				using (var stream = new MemoryStream())
				{
					workbook.SaveAs(stream);
					return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExportExcel.xlsx");
				}
			}
		}





		public ActionResult Edit(int id)
		{
			DataTable dt = GetExcelDataById(id); // Veritabanından veriyi çek (DataTable olarak)
			if (dt.Rows.Count > 0)
			{
				return View(dt.Rows[0]); // İlk satırı View'a gönder
			}
			return RedirectToAction("Index"); // Kayıt bulunamadıysa ana sayfaya yönlendir
		}

		// Düzenleme işlemini gerçekleştir
		[HttpPost]
		public ActionResult Edit(FormCollection form)
		{
			if (ModelState.IsValid)
			{
				UpdateExcelData(form); // FormCollection'dan verileri alarak güncelle
				return RedirectToAction("Index");
			}

			// Model state geçersizse, veriyi tekrar çek ve View'a gönder
			DataTable dt = GetExcelDataById(Convert.ToInt32(form["Id"]));
			return View(dt.Rows[0]);
		}

		// Veritabanından ID'ye göre veri çekme (DataTable olarak)
		private DataTable GetExcelDataById(int id)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				string query = "SELECT * FROM DataTable WHERE Id = @Id";
				using (SqlCommand command = new SqlCommand(query, connection))
				{
					command.Parameters.AddWithValue("@Id", id);
					SqlDataAdapter adapter = new SqlDataAdapter(command);
					DataTable dataTable = new DataTable();
					adapter.Fill(dataTable);
					return dataTable;
				}
			}
		}

		// Veritabanında güncelleme (DataTable kullanarak)
		private void UpdateExcelData(FormCollection form)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				string query = "UPDATE DataTable SET Ad = @Ad, Soyad = @Soyad, Yas = @Yas, Sehir = @Sehir, Tarih = @Tarih WHERE Id = @Id";

				using (SqlCommand command = new SqlCommand(query, connection))
				{
					// FormCollection'dan gelen verileri parametre olarak ekle
					command.Parameters.AddWithValue("@Ad", form["Ad"]);
					command.Parameters.AddWithValue("@Soyad", form["Soyad"]);
					command.Parameters.AddWithValue("@Yas", Convert.ToInt32(form["Yas"]));
					command.Parameters.AddWithValue("@Sehir", form["Sehir"]);
					command.Parameters.AddWithValue("@Tarih", Convert.ToDateTime(form["Tarih"]));
					command.Parameters.AddWithValue("@Id", Convert.ToInt32(form["Id"]));
					command.ExecuteNonQuery();
				}
			}
		}




























		//Silme İşlemi
		[HttpPost]
		public ActionResult Sil(int id)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand("DELETE FROM DataTable WHERE Id = @Id", connection);
				command.Parameters.AddWithValue("@Id", id);
				connection.Open();
				command.ExecuteNonQuery();
			}

			return RedirectToAction("Index");
		}
	}
}