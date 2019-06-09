using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Volby.Model;

namespace Volby
{
    class Program
    {
        static void Main(string[] args)
        {
            var scraper = new DistrictScraper();
            scraper.ScrapeMunicipality();
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();
        }
    }
}
