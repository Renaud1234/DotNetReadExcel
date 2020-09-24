using ExcelWriter.Models;
using ExcelWriter.Tools;
using Microsoft.AspNetCore.Http;
using System.IO;
using Xunit;

namespace XUnitExcelWriter
{
    public class SerialiserShould
    {
        [Fact]
        public void DeserializeReportsXml()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\TestFiles\Test.xml");
            FileStream xmlFile = File.OpenRead(path);
            var ms = new MemoryStream();
            xmlFile.CopyTo(ms);
            IFormFile formFile = new FormFile(ms, 0, ms.Length, path, "Test.xml");

            Reports reports = Serializer.Deserialize<Reports>(formFile);

            Assert.Equal("F 20.04", reports.Report.Name);
            Assert.Equal(6, reports.Report.ReportVal.Count);
            Assert.Equal("200", reports.Report.ReportVal[1].Val);
        }
    }
}
