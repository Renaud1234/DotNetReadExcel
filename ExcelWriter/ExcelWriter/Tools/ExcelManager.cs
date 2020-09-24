using ExcelWriter.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Threading.Tasks;

namespace ExcelWriter.Tools
{
    public class ExcelManager : IExcelManager
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public ExcelManager(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public async Task<MemoryStream> Merge(Reports reports, IFormFile excelFile)
        {
            string sFileName = excelFile.FileName;
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            XSSFWorkbook workbook;
            MemoryStream memory = new MemoryStream();

            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.ReadWrite))
            {
                excelFile.CopyTo(stream);
                stream.Position = 0;
                workbook = new XSSFWorkbook(stream);
            }

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

            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.ReadWrite))
            {
                workbook.Write(stream);
            }

            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }

            memory.Position = 0;
            return memory;
        }
    }
}
