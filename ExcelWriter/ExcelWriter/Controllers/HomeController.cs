using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ExcelWriter.Models;
using ExcelWriter.Tools;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace ExcelWriter.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IExcelManager _excelManager;

        public HomeController(ILogger<HomeController> logger, IExcelManager excelManager)
        {
            _logger = logger;
            _excelManager = excelManager;
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
                
                if (excelFile.Length > 0)
                {

                    //FileStreamResult outputFile = Merge(reports, excelFile);
                    MemoryStream memory = await _excelManager.Merge(reports, excelFile);

                    return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelFile.FileName);
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
                ViewData["ErrorMessage"] = "Error Excel file.";
                return View();
            }            
        }
    }
}
