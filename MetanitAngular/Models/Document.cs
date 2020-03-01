using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using Microsoft.AspNetCore.Http;

namespace MetanitAngular.Models
{
    public class Document
    {
        public IFormFile file { get; set; }
        public string name { get; set; }
    }
}
