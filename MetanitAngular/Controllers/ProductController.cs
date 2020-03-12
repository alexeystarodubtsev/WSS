using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using MetanitAngular.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using MetanitAngular.Parsers;
using System.Collections.ObjectModel;
using System;

namespace MetanitAngular.Controllers
{
    [ApiController]
    [Route("api/products")]
    public class ProductController : Controller
    {
        ApplicationContext db;
        public ProductController(ApplicationContext context)
        {
            db = context;

            
            if (!db.Products.Any())
            {
                db.Products.Add(new Product { Name = "iPhone X", Company = "Apple", Price = 79900 });
                db.Products.Add(new Product { Name = "Galaxy S8", Company = "Samsung", Price = 49900 });
                db.Products.Add(new Product { Name = "Pixel 2", Company = "Google", Price = 52900 });
                db.SaveChanges();
            }
        }
        [HttpGet]
        public IEnumerable<Product> Get()
        {
            
            Product p = new Product();
            p.Id = 1;
            p.Name = "sjs";
            p.Price = 30;
            p.Company = "ss";
            List<Product> lp = new List<Product>();
            lp.Add(p);
            lp = db.Products.ToList();
            return lp;
        }

        [HttpGet("{id}")]
        public Product Get(int id)
        {
            Product product = db.Products.FirstOrDefault(x => x.Id == id);
            return product;
        }

        //[HttpPost]
        //public IActionResult Post(Product product)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        db.Products.Add(product);
        //        db.SaveChanges();
        //        return Ok(product);
        //    }
        //    return BadRequest(ModelState);
        //}
        [HttpPost]
        public Tuple<List<Phone>, List<Phone>> Post()
        {
            return XlPhone.getPhone(Request.Form.Files);
        }
        [HttpPut]
        public IActionResult Put(Product product)
        {
            if (ModelState.IsValid)
            {
                db.Update(product);
                db.SaveChanges();
                return Ok(product);
            }
            return BadRequest(ModelState);
        }

        [HttpDelete("{id}")]
        public IActionResult Delete(int id)
        {
            Product product = db.Products.FirstOrDefault(x => x.Id == id);
            if (product != null)
            {
                db.Products.Remove(product);
                db.SaveChanges();
            }
            return Ok(product);
        }
    }
}