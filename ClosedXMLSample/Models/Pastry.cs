using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ClosedXMLSample.Models
{
    public class Pastry
    {
        public Pastry(string name, int amount, string month)
        {
            Month = month;
            Name = name;
            NumberOfOrders = amount;
        }

        public string Name { get; set; }
        public int NumberOfOrders { get; set; }
        public string Month { get; set; }
    }
}