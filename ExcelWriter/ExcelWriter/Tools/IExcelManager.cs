using ExcelWriter.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;

namespace ExcelWriter.Tools
{
    public interface IExcelManager
    {
        Task<MemoryStream> Merge(Reports reports, IFormFile excelFile);
    }
}