using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelApi.Models
{
    public class ExcelEntity
    {
        public string Sno { get; set; }
        public string Company { get; set; }
        public string Sector { get; set; }
        public string SubSector { get; set; }
        public string Region { get; set; }
        public string NoOfEmployee { get; set; }
        public string TotalRevenu { get; set; }
        public string Websites { get; set; }
    }
}