namespace Volby;

class Program
{
    static void Main(string[] args)
    {
        var isDebugMode = (args?.Any(item => item == "--debug") == true);
        var scraper = new DistrictScraper(isDebugMode);
        scraper.ScrapeMunicipality();
        Console.WriteLine("Press enter to continue");
        Console.ReadLine();
    }
}
