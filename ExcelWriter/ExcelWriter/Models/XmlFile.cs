using Newtonsoft.Json;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExcelWriter.Models
{
    public partial class XmlFile
    {
        [JsonProperty("Reports")]
        public Reports Reports { get; set; }
    }

    [XmlRoot(ElementName = "Reports")]
    public partial class Reports
    {
        [XmlElement(ElementName = "Report")]
        public Report Report { get; set; }
    }

    [XmlRoot(ElementName = "Report")]
    public class Report
    {
        [XmlElement(ElementName = "Name")]
        public string Name { get; set; }
        [XmlElement(ElementName = "ReportVal")]
        public List<ReportVal> ReportVal { get; set; }
    }

    [XmlRoot(ElementName = "ReportVal")]
    public class ReportVal
    {
        [XmlElement(ElementName = "ReportRow")]
        public string ReportRow { get; set; }
        [XmlElement(ElementName = "ReportCol")]
        public string ReportCol { get; set; }
        [XmlElement(ElementName = "Val")]
        public string Val { get; set; }
    }
}
