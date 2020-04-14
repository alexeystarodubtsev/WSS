using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class DataCall
    {
        public string Stage { get; set; }
        public DateTime Date { get; set; }
        public DataCall(string stage, DateTime date)
        {
            Stage = stage;
            Date = date;
        }
    }
}
