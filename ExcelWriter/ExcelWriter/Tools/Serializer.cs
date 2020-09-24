using Microsoft.AspNetCore.Http;
using System.IO;
using System.Xml.Serialization;

namespace ExcelWriter.Tools
{
    public class Serializer
    {
        public static T Deserialize<T>(IFormFile xmlFile) where T : class
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using StreamReader reader = new StreamReader(xmlFile.OpenReadStream());
            return (T)serializer.Deserialize(reader);
        }
    }
}
