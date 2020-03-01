using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }    // название 
        public string Company { get; set; } // производитель
        public int Price { get; set; }  // цена
    }
}
