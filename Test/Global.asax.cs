using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace Test
{
	public class MvcApplication : System.Web.HttpApplication
	{
		protected void Application_Start(object sender , EventArgs e)
		{
			//string connectionString = ConfigurationManager.ConnectionStrings["Context"].ConnectionString;

			//if (string.IsNullOrEmpty(connectionString))
			//{
			//	// Yapılandırma sayfasına yönlendir
			//	Response.Redirect("~/DatabaseConfig.aspx");
			//}
			//else
			//{
			//	// Bağlantı bilgileri mevcut, veritabanı kontrolüne geç
			//	CheckAndConfigureDatabase(connectionString);
			//}
			AreaRegistration.RegisterAllAreas();
			FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
			RouteConfig.RegisterRoutes(RouteTable.Routes);
			BundleConfig.RegisterBundles(BundleTable.Bundles);
		}

		private void CheckAndConfigureDatabase(string connectionString)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				try
				{
					connection.Open();
					// Veritabanı var, temizleme seçeneği sun
					// ... (Kullanıcıya onay sorma ve tabloları silme işlemleri)
				}
				catch (SqlException ex)
				{
					if (ex.Number == 4060) // Veritabanı bulunamadı hatası
					{
						// Veritabanını oluştur
						CreateDatabase(connectionString);
					}
					else
					{
						// Diğer hataları yönet
						// ...
					}
				}
			}
		}

		private void CreateDatabase(string connectionString)
		{
			// Bağlantı dizesini ayrıştır
			SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectionString);

			// Veritabanı adını al
			string databaseName = builder.InitialCatalog;

			// Veritabanı adını bağlantı dizesinden kaldır (master veritabanına bağlanmak için)
			builder.InitialCatalog = "master";

			using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
			{
				connection.Open();

				// Veritabanı oluşturma komutu
				string createDbQuery = $"CREATE DATABASE [{databaseName}]";

				using (SqlCommand command = new SqlCommand(createDbQuery, connection))
				{
					command.ExecuteNonQuery();
				}
			}
		}
	}
}
