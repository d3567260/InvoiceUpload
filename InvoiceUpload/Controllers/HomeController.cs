using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using InvoiceUpload.Dtos;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using InvoiceUpload.Models;
using Microsoft.AspNetCore.Http;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace InvoiceUpload.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly string _token;
        private readonly string _time;
        private readonly bool _isLogin;

        public HomeController(ILogger<HomeController> logger, IHttpContextAccessor httpContextAccessor)
        {
            _logger = logger;
            _token = httpContextAccessor.HttpContext?.Session.GetString("_Token");
            _time = httpContextAccessor.HttpContext?.Session.GetString("_Time");
            _isLogin = _token is not null && _time is not null &&
                       (DateTime.Now - DateTime.Parse(_time)).Duration() < TimeSpan.FromMinutes(600);
        }

        public IActionResult Index()
        {
            if (_isLogin)
            {
                return RedirectToAction(nameof(UploadExcel));
            }

            return View();
        }

        [HttpPost("login")]
        [ValidateAntiForgeryToken]
        public IActionResult Login(LoginDto loginDto)
        {
            if (ModelState.IsValid)
            {
                // Create a request using a URL that can receive a post.
                WebRequest request = WebRequest.Create("https://middleware.tainanshopping.tw/api/member/token");
                // Set the Method property of the request to POST.
                request.Method = "POST";

                // Create POST data and convert it to a byte array.
                var postData = new Dictionary<string, string>
                {
                    { "cellphone", loginDto.Phone.Remove(0, 1) },
                    { "password", loginDto.Password },
                    { "country_code", "+886" },
                    { "country", "TW" }
                };
                byte[] byteArray = JsonSerializer.SerializeToUtf8Bytes(postData);

                // Set the ContentType property of the WebRequest.
                request.ContentType = "application/json";
                // Set the ContentLength property of the WebRequest.
                request.ContentLength = byteArray.Length;

                // Get the request stream.
                Stream dataStream = request.GetRequestStream();
                // Write the data to the request stream.
                dataStream.Write(byteArray, 0, byteArray.Length);
                // Close the Stream object.
                dataStream.Close();

                // Get the response.
                WebResponse response = request.GetResponse();
                // Display the status.
                _logger.LogInformation(((HttpWebResponse)response).StatusDescription);

                // Get the stream containing content returned by the server.
                // The using block ensures the stream is automatically closed.
                using (dataStream = response.GetResponseStream())
                {
                    // Open the stream using a StreamReader for easy access.
                    StreamReader reader = new StreamReader(dataStream);
                    // Read the content.
                    string responseFromServer = reader.ReadToEnd();
                    _logger.LogInformation(responseFromServer);
                    JObject json = JObject.Parse(responseFromServer);
                    
                    if ((string)json["code"] == "00000")
                    {
                        HttpContext.Session.SetString("_Token", (string)json["data"]?["token"]);
                        HttpContext.Session.SetString("_Time", DateTime.Now.ToString());
                    }
                    else
                    {
                        TempData["ErrorMessage"] = "登入失敗，請重新登入。";
                        return RedirectToAction(nameof(Index));
                    }
                }

                // Close the response.
                response.Close();
            }

            return RedirectToAction(nameof(UploadExcel));
        }

        [HttpGet("upload-excel")]
        public IActionResult UploadExcel()
        {
            if (!_isLogin)
            {
                TempData["ErrorMessage"] = "尚未登入，請先登入。";
                return RedirectToAction(nameof(Index));
            }

            return View();
        }

        [HttpPost("upload-excel")]
        [ValidateAntiForgeryToken]
        public IActionResult UploadExcelAction(UploadExcelDto uploadExcelDto)
        {
            if (ModelState.IsValid)
            {
                if (_isLogin)
                {
                    using var memoryStream = new MemoryStream();
                    uploadExcelDto.FormFile.CopyTo(memoryStream);

                    // Upload the file if less than 2 MB
                    if (memoryStream.Length < 2097152)
                    {
                        using var workBook = new XLWorkbook(memoryStream);

                        var workSheet = workBook.Worksheet(1);

                        var data = workSheet.Range($"A2:{workSheet.LastCellUsed()}");

                        foreach (var row in data.RowsUsed())
                        {
                            // Create a request using a URL that can receive a post.
                            WebRequest request =
                                WebRequest.Create("https://middleware.tainanshopping.tw/api/v1/invoice/addByPaper");
                            // Set the Method property of the request to POST.
                            request.Method = "POST";
                            request.Headers.Add("Authorization", "Bearer " + _token);

                            // Create POST data and convert it to a byte array.
                            var postData = new Dictionary<string, string>
                            {
                                { "invoiceDateY", row.Cell(1).GetValue<string>() },
                                { "invoiceDateM", row.Cell(2).GetValue<string>() },
                                { "invoiceDateD", row.Cell(3).GetValue<string>() },
                                { "invoiceTotal", row.Cell(4).GetValue<string>() },
                                { "invoiceNo", row.Cell(5).GetValue<string>() },
                                { "companyNo", row.Cell(6).GetValue<string>() }
                            };

                            _logger.LogInformation(JsonSerializer.Serialize(postData));

                            byte[] byteArray = JsonSerializer.SerializeToUtf8Bytes(postData);

                            // Set the ContentType property of the WebRequest.
                            request.ContentType = "application/json";
                            // Set the ContentLength property of the WebRequest.
                            request.ContentLength = byteArray.Length;

                            // Get the request stream.
                            Stream dataStream = request.GetRequestStream();
                            // Write the data to the request stream.
                            dataStream.Write(byteArray, 0, byteArray.Length);
                            // Close the Stream object.
                            dataStream.Close();

                            // Get the response.
                            WebResponse response = request.GetResponse();
                            // Display the status.
                            _logger.LogInformation(((HttpWebResponse)response).StatusDescription);

                            // Get the stream containing content returned by the server.
                            // The using block ensures the stream is automatically closed.
                            using (dataStream = response.GetResponseStream())
                            {
                                // Open the stream using a StreamReader for easy access.
                                StreamReader reader = new StreamReader(dataStream);
                                // Read the content.
                                string responseFromServer = reader.ReadToEnd();
                                _logger.LogInformation(responseFromServer);
                            }

                            // Close the response.
                            response.Close();
                        }
                    }
                    else
                    {
                        TempData["ErrorMessage"] = "檔案大小超過2MB。";
                    }
                }
                else
                {
                    TempData["ErrorMessage"] = "尚未登入，請先登入。";
                    return RedirectToAction(nameof(Index));
                }

                TempData["SuccessMessage"] = "匯入成功。";
            }
            else
            {
                TempData["ErrorMessage"] = "匯入失敗。";    
            }

            return RedirectToAction(nameof(UploadExcel));
        }

        [HttpGet("download-excel")]
        public IActionResult DownloadExcel()
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("工作表一");
                    worksheet.Cell("A1").Value = "發票年";
                    worksheet.Cell("B1").Value = "發票月";
                    worksheet.Cell("C1").Value = "發票日";
                    worksheet.Cell("D1").Value = "發票金額";
                    worksheet.Cell("E1").Value = "發票號碼";
                    worksheet.Cell("F1").Value = "公司統編";

                    using (var memoryStream = new MemoryStream())
                    {
                        workbook.SaveAs(memoryStream);
                        return File(memoryStream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "發票匯入檔.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return Error();
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}