using Dapper;
using ImportData.Data;
using ImportData.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace ImportData.Controllers
{
    //[ApiController]
    //[Route("api/[controller]")]
    public class HomeController : Controller
    {
        private readonly Microsoft.Extensions.Configuration.IConfiguration config;

        

        private readonly CountriesAPIDbContext dbContext;

        public HomeController(CountriesAPIDbContext dbContext, Microsoft.Extensions.Configuration.IConfiguration _config)
        {
            this.dbContext = dbContext;
            this.config=_config;
    }

        [HttpGet]
        public async Task<IActionResult> GetCountries()
        {
            var countries = await dbContext.Countries.ToListAsync();
            return Ok(countries);
        }

        //[HttpGet]
        //[Route("ExportToExcel")]
       public async Task<IActionResult> ExportToExcel()
{
    try
    {
        var data = await dbContext.Countries.ToListAsync();

        // Create a new Excel package
        using (var package = new ExcelPackage())
        {
            // Add a new worksheet to the Excel package
            var worksheet = package.Workbook.Worksheets.Add("Countries");

            // Set the column headers
            worksheet.Cells[1, 1].Value = "Country ID";
            worksheet.Cells[1, 2].Value = "Country Name";
            worksheet.Cells[1, 3].Value = "Two Char Country Code";
            worksheet.Cells[1, 4].Value = "Three Char Country Code";

            int row = 2;
            foreach (var country in data)
            {
                worksheet.Cells[row, 1].Value = country.CountryID;
                worksheet.Cells[row, 2].Value = country.CountryName;
                worksheet.Cells[row, 3].Value = country.TwoCharCountryCode;
                worksheet.Cells[row, 4].Value = country.ThreeCharCountryCode;
                row++;
            }

            // Export the Excel file and return it as the response
            // Auto-fit the columns
            worksheet.Cells.AutoFitColumns();

            // Convert the Excel package to a byte array
            var excelBytes = package.GetAsByteArray();

            // Set the response content type and headers
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.Headers.Add("Content-Disposition", "attachment; filename=Countries.xlsx");

            // Write the Excel data to the response stream
            await Response.Body.WriteAsync(excelBytes);

            return Ok();
        }
    }
    catch (Exception ex)
    {
        // Handle the exception here, you can log the error or return a specific error response
        return StatusCode(StatusCodes.Status500InternalServerError, "An error occurred during Excel export: " + ex.Message);
    }
}


        public IActionResult Index()
        {
            ViewData["Title"] = "Home Page";
            return View();
        }



        [HttpPost("ConvertToJSON")]
        public async Task<IActionResult> Import(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("File Not Found.");
                }

                var fileExtension = Path.GetExtension(file.FileName).ToLower();
                if (fileExtension != ".xlsx")
                {
                    return BadRequest("Invalid File");
                }

                // Read the Excel file into a DataTable
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    stream.Position = 0;

                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var dataTable = new DataTable();

                        foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                        {
                            dataTable.Columns.Add(firstRowCell.Text);
                        }

                        for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                        {
                            var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                            var newRow = dataTable.NewRow();
                            foreach (var cell in row)
                            {
                                newRow[cell.Start.Column - 1] = cell.Text;
                            }
                            dataTable.Rows.Add(newRow);
                        }

                        // Call the stored procedure to insert data into the database
                        string dbConn = config.GetSection("ConnectionStrings").GetSection("CountriesAPIConnectionString").Value;
                        using (var connection = new SqlConnection(dbConn))
                        {
                            connection.Open();

                            var parameters = new DynamicParameters();
                            parameters.Add("ParCountryType", dataTable.AsTableValuedParameter("dbo.NewCountryType"));

                            var sqlQuery = "Usp_ValidationInsertCountries  "; 

                            await connection.ExecuteAsync(sqlQuery, parameters, commandType: CommandType.StoredProcedure);
                        }
                    }
                }

                // Return the success message
                return Ok("Data inserted into the database from Excel file.");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error: {ex.Message}");
            }
        }



       // [HttpPost("ConvertToJSON")]
        //public async Task<IActionResult> Import(IFormFile file)
        //{
        //    try
        //    {
        //        if (file == null || file.Length == 0)
        //        {
        //            return BadRequest("File Not Found.");
        //        }

        //        var fileExtension = Path.GetExtension(file.FileName).ToLower();
        //        if (fileExtension != ".csv")
        //        {
        //            return BadRequest("Invalid File");
        //        }

        //        // Create the destination file path with .json extension
        //        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);
        //        var destFileName = $"{fileNameWithoutExtension}.json";
        //        var destPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, destFileName);


        //        using (var stream = new MemoryStream())
        //        {
        //            await file.CopyToAsync(stream);
        //            stream.Position = 0;

        //            using (var reader = new StreamReader(stream))
        //            using (var csv = new CsvHelper.CsvReader(reader, new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)))
        //            {
        //                // Read the CSV data and deserialize it into a list of Countries objects
        //                var records = csv.GetRecords<Countries>().ToList();

        //                // Serialize the list of Countries objects into JSON
        //                var json = JsonSerializer.Serialize(records);

        //                // Write the JSON data to the destination file
        //                await System.IO.File.WriteAllTextAsync(destPath, json);

        //                // Read the generated JSON data
        //                var jsonData = await System.IO.File.ReadAllTextAsync(destPath);

        //                // Call the stored procedure to insert data into the database
        //                using (var connection = new SqlConnection("CountriesAPIConnectionString"))
        //                {
        //                    connection.Open();

        //                    var parameters = new
        //                    {
        //                        JsonData = jsonData
        //                    };

        //                    var sqlQuery = "EXEC dbo.InsertCountriesFromJson @JsonData"; // Modify this if the stored procedure call is different.

        //                    await connection.ExecuteAsync(sqlQuery, parameters);
        //                }

        //                // Return the success message with the file path
        //                return Ok($"Data inserted into the database from JSON file: {destPath}");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Error: {ex.Message}");
        //    }
        //}












        //private List<Countries> countries = new List<Countries>
        //{
        //    new Countries{ CountryID="11" ,CountryName="India", TwoCharCountryCode="IN",ThreeCharCountryCode="IND"},
        //    new Countries{ CountryID="12" ,CountryName="Australia", TwoCharCountryCode="AU",ThreeCharCountryCode="AUS"}
        //};


        //public IActionResult ExportToCSV(Countries countries)
        //{
        //    var builder = new StringBuilder();
        //    builder.AppendLine(" CountryID ,CountryName, TwoCharCountryCode,ThreeCharCountryCode");
        //    foreach (var country in Countries)
        //    {
        //        builder.AppendLine($"{country.CountryID},{country.CountryName},{country.TwoCharCountryCode},{country.ThreeCharCountryCode}");
        //    }
        //    return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "countries.csv");
        //}
        //public IActionResult Index()
        //{
        //    return View();
        //}


        //[HttpPost("ImportCountries")]
        //public async Task<IActionResult> Import(IFormFile file)
        //{
        //    try
        //    {
        //        var countries = await ImportfromExcel(file);
        //        dbContext.Countries.AddRange(countries);
        //        await dbContext.SaveChangesAsync();

        //        return Ok("Data imported successfully");
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Error: {ex.Message}");
        //    }
        //}





        //STORED PROCEDURE
        //    private readonly string connectionString = "YourConnectionStringHere";

        //[HttpPost("ImportCountries")]
        //public IActionResult Import(IFormFile file)
        //{
        //    try
        //    {
        //        var filePath = Path.GetTempFileName();
        //        using (var stream = new FileStream(filePath, FileMode.Create))
        //        {
        //            file.CopyTo(stream);
        //        }
        //        // Read the Excel file
        //        var countries = ImportfromExcel(filePath);

        //        // Check for duplicate CountryID values
        //        var duplicateCountryIDs = countries.GroupBy(c => c.CountryID)
        //                                           .Where(g => g.Count() > 1)
        //                                           .Select(g => g.Key)
        //                                           .ToList();

        //        if (duplicateCountryIDs.Any())
        //        {
        //            var duplicateCountryIDsString = string.Join(", ", duplicateCountryIDs);
        //            return BadRequest($"Duplicate CountryID values found: {duplicateCountryIDsString}");
        //        }

        //        // Import the data to the database
        //        dbContext.Countries.AddRange(countries);
        //        dbContext.SaveChanges();




        //        return Ok("Data imported successfully");
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Error: {ex.Message}");
        //    }
        //}


        //[HttpPost]
        //[Route("ImportfromExcel")]

        //public IActionResult Import(IFormFile file)
        //{
        //    // Check if the file is not null and is a valid Excel file
        //    if (file != null && file.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //    {
        //        try
        //        {
        //            // Save the uploaded file to a temporary location
        //            var filePath = Path.GetTempFileName();
        //            using (var stream = new FileStream(filePath, FileMode.Create))
        //            {
        //                file.CopyTo(stream);
        //            }

        //            // Call the Import method with the valid file path
        //            UploadFile(filePath);

        //            return Ok("Data imported successfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            return BadRequest($"Error: {ex.Message}");
        //        }
        //    }
        //    else
        //    {
        //        return BadRequest("Invalid file. Please upload an Excel file.");
        //    }
        //}






        //public void UploadFile(string filePath)
        //{
        //    using (var package = new ExcelPackage(new FileInfo(filePath)))
        //    {
        //        var worksheet = package.Workbook.Worksheets[0];
        //        var rowCount = worksheet.Dimension.Rows;

        //        using (var connection = new SqlConnection("CountriesAPIConnectionString"))
        //        {
        //            connection.Open();

        //            using (var command = new SqlCommand("dbo.ImportDataFromExcel", connection))
        //            {
        //                command.CommandType = CommandType.StoredProcedure;

        //                for (int row = 2; row <= rowCount; row++)
        //                {
        //                    command.Parameters.Clear();
        //                    command.Parameters.AddWithValue("@CountryID", worksheet.Cells[row, 1].Value?.ToString()?.Trim());
        //                    command.Parameters.AddWithValue("@CountryName", worksheet.Cells[row, 2].Value?.ToString()?.Trim());
        //                    command.Parameters.AddWithValue("@TwoCharCountryCode", worksheet.Cells[row, 3].Value?.ToString()?.Trim());
        //                    command.Parameters.AddWithValue("@ThreeCharCountryCode", worksheet.Cells[row, 4].Value?.ToString()?.Trim());

        //                    command.ExecuteNonQuery();
        //                }
        //            }

        //            connection.Close();
        //        }
        //    }
        //}
        //public async Task<List<Countries>> ImportfromExcel(IFormFile file)
        //{
        //    var list = new List<Countries>();
        //    using (MemoryStream stream = new MemoryStream())

        //    {
        //        await file.CopyToAsync(stream);
        //        using (var package = new ExcelPackage(stream))
        //        {

        //            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        //            var rowcount = worksheet.Dimension.Rows;
        //            if (rowcount > 0)
        //            {
        //                for (int row = 2; row <= rowcount; row++)
        //                {
        //                    list.Add(new Countries
        //                    {
        //                        CountryID = worksheet.Cells[row, 1].Value.ToString().Trim(),
        //                        CountryName = worksheet.Cells[row, 2].Value.ToString().Trim(),
        //                        TwoCharCountryCode = worksheet.Cells[row, 3].Value.ToString().Trim(),
        //                        ThreeCharCountryCode = worksheet.Cells[row, 4].Value.ToString().Trim()
        //                    });
        //                }
        //            }
        //            else
        //            {

        //                Console.WriteLine("The Excel sheet is empty.");

        //            }


        //        }



        //    }
        //    return list;


        //}

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