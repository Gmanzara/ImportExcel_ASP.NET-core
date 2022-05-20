using ExcelImport.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;


namespace ExcelImport.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }


        [HttpPost]
        public async Task<IActionResult> ImportExcelFile(IFormFile FormFile)
        {
            //obtenir le nom du fichier
            var filename = ContentDispositionHeaderValue.Parse(FormFile.ContentDisposition).FileName.Trim('"');

            //obtenir le chemin
            var MainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");

            //créer le repectoire "Uploads" s'il n'existe pas
            if (!Directory.Exists(MainPath))
            {
                Directory.CreateDirectory(MainPath);
            }

            //obtenir le chemin du fichier 
            var filePath = Path.Combine(MainPath, FormFile.FileName);
            using (System.IO.Stream stream = new FileStream(filePath, FileMode.Create))
            {
                await FormFile.CopyToAsync(stream);
            }

            //obtenir l'extension
            string extension = Path.GetExtension(filename);


            string conString = string.Empty;

            switch (extension)
            {
                case ".xls": //Excel 97-03.
                    conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                    break;
                case ".xlsx": //Excel 07 and above.
                    conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                    break;
            }

            DataTable dt = new DataTable();
            conString = string.Format(conString, filePath);

            using (OleDbConnection connExcel = new OleDbConnection(conString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        cmdExcel.Connection = connExcel;

                        //obtenir le nom de la première feuille.
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        connExcel.Close();

                        //Lire les données de la prémière feuille.
                        connExcel.Open();
                        cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dt);
                        connExcel.Close();
                    }
                }
            }
            //votre chaîne de connexion à la base de données
            conString = @"Data Source=CNTIG-PC302;Initial Catalog=BD_ImportDonnes;Integrated Security=False;Persist Security Info=False;User ID=sa;Password=Cntig2022";

            using (SqlConnection con = new SqlConnection(conString))
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    //Définir le nom de la table de la base de données.
                    sqlBulkCopy.DestinationTableName = "dbo.Personne";

                    // Mettez en correspondance les colonnes d'Excel avec celles de la table de la base de données, ceci est facultatif mais utile si vous le faites.

                    // [FACULTATIF] : Mettez en correspondance les colonnes Excel avec celles de la table de la base de données.
                    sqlBulkCopy.ColumnMappings.Add("id", "id");
                    sqlBulkCopy.ColumnMappings.Add("nom", "nom");
                    sqlBulkCopy.ColumnMappings.Add("prenom", "prenom");
                    sqlBulkCopy.ColumnMappings.Add("nationalite", "nationalite");

                    con.Open();
                    sqlBulkCopy.WriteToServer(dt);
                    con.Close();
                }
            }
            //si le code atteint ici signifie que tout va bien et que les données Excel sont importées dans la base de données.
            ViewBag.Message = "Fichier importé et données excel enregistrées dans la base de données";

            return View("Index");

        }


        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
