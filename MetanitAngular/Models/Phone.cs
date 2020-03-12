using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class Phone
    {
        public string PhoneNumber { get; set; }
        public int Qty { get; set; }
        public string Stage { get; set; }
        public string Date { get; set; }
        public Phone(string phone, int qty, string stage, string date)
        {
            PhoneNumber = phone;
            Qty = qty;
            Stage = stage;
            Date = date;
        }
    }
}
