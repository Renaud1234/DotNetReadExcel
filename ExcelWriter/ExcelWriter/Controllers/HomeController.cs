using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelWriter.Models;
using ExcelWriter.Tools;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;

namespace ExcelWriter.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("UploadFiles")]
        public async Task<IActionResult> Index(List<IFormFile> files)
        {
            if (files.Count != 2)
            {
                ViewData["ErrorMessage"] = "Select XML and Excel files.";
                return View();
            }

            Reports reports = new Reports();
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName;
            MemoryStream memory = new MemoryStream();

            try
            {   // Read XML file
                IFormFile xmlFile = files[0];
                if (xmlFile.Length > 0)
                {                    
                    reports = Serializer.Deserialize<Reports>(xmlFile);
                }
                else
                {
                    ViewData["ErrorMessage"] = "XML file is empty.";
                    return View();
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex.ToString());
                ViewData["ErrorMessage"] = "Error XML file reading.";
            }

            try
            {   // Update Excel file
                IFormFile excelFile = files[1];
                XSSFWorkbook workbook;
                if (excelFile.Length > 0)
                {
                    sFileName = excelFile.FileName;
                    using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.ReadWrite))
                    {
                        excelFile.CopyTo(stream);
                        stream.Position = 0;
                        workbook = new XSSFWorkbook(stream);

                        var finalSheet = workbook.GetSheet("Final");

                        foreach (var reportVal in reports.Report.ReportVal)
                        {
                            int cellRow = 0;
                            int cellCol = 0;

                            for (int row = 10; row <= 11; row++)
                            {
                                if (finalSheet.GetRow(row).GetCell(1).StringCellValue.Substring(1) == reportVal.ReportRow)
                                {
                                    cellRow = row;
                                    break;
                                }
                            }
                            for (int col = 4; col <= 10; col++)
                            {
                                if (finalSheet.GetRow(9).GetCell(col).StringCellValue.Substring(1) == reportVal.ReportCol)
                                {
                                    cellCol = col;
                                    break;
                                }
                            }
                            if (cellRow != 0 && cellCol != 0)
                            {
                                var cell = finalSheet.GetRow(cellRow).GetCell(cellCol);
                                cell.SetCellValue(reportVal.Val);
                            }
                        }
                    }
                    using(var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open, FileAccess.ReadWrite))
                    {
                        workbook.Write(stream);
                    }
                } 
                else
                {
                    ViewData["ErrorMessage"] = "Excel file is empty.";
                    return View();
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex.ToString());
                ViewData["ErrorMessage"] = "Error Excel file" + ex.ToString();
                return View();
            }

            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName);
        }
    }
}
