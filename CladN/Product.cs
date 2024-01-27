using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CladN
{
    internal class Product
    {
        public int Code { get; set; }    
        public string Name { get; set; } 
        public double Price { get; set; }  
        public string Unit { get; set; }

        public readonly int Id;

        public Product(int id, int code, string name, string unit, double price) 
        {
            Id = id;
            Code = code;
            Name = name;
            Unit = unit;
            Price = price;
        }

        public void SetCode(int code)
        {
            Code = code;
        }
        public void SetName(string name)
        {
            Name = name;
        }
        public void SetUnit(string unit)
        {
            Unit = unit;
        }
        public void SetPrice(double price)
        {
            Price = price;
        }
    }
}
