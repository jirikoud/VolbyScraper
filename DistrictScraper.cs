using HtmlAgilityPack;
using NPOI.HSSF.UserModel;
using System.Web;
using Volby.Model;

namespace Volby;

public class DistrictScraper(
    bool isDebugMode
    ) : IDisposable
{
    private const string BASE_URL = "https://www.volby.cz/pls/kv2022/";

    private bool _isDebugMode = isDebugMode;
    private HttpClient _httpClient = new HttpClient();

    public void Dispose()
    {
        _httpClient.Dispose();
    }

    private int GetIntValue(HtmlNode node)
    {
        var stringValue = node.InnerText;
        stringValue = stringValue.Replace("&nbsp;", "");
        return int.Parse(stringValue);
    }

    private int GetCellIntValue(HtmlDocument document, string xPath)
    {
        var stringValue = document.DocumentNode.SelectSingleNode(xPath).InnerText;
        stringValue = stringValue.Replace("&nbsp;", "");
        return int.Parse(stringValue);
    }

    private double GetDoubleValue(HtmlNode node)
    {
        var stringValue = node.InnerText;
        stringValue = stringValue.Replace("&nbsp;", "");
        stringValue = stringValue.Replace(",", ".");
        return double.Parse(stringValue);
    }

    private HtmlDocument LoadHtmlDocument(string url)
    {
        string htmlCode = _httpClient.GetStringAsync(url).Result;
        var document = new HtmlDocument();
        document.LoadHtml(htmlCode);
        return document;
    }

    private void LoadParties(MunicipalityModel municipality)
    {
        try
        {
            var document = LoadHtmlDocument(municipality.Url);

            municipality.UsedEnvelopes = GetCellIntValue(document, "//td[@class='cislo'][@headers='sa7']");
            municipality.ValidVotes = GetCellIntValue(document, "//td[@class='cislo'][@headers='sa8']");

            var nodeList = document.DocumentNode.SelectNodes("//td[@class='cislo'][@headers='t2sa2 t2sb1']");
            if (nodeList != null)
            {
                foreach (var node in nodeList)
                {
                    var parent = node.ParentNode;
                    if (node.FirstChild == null || !node.FirstChild.HasAttributes)
                    {
                        continue;
                    }
                    var urlPath = node.FirstChild.Attributes["href"].Value;
                    var nameNode = parent.SelectSingleNode("td[@headers='t2sa2 t2sb2']");
                    var party = new PartyModel()
                    {
                        Url = BASE_URL + HttpUtility.HtmlDecode(urlPath),
                        Name = nameNode.InnerText,
                    };
                    municipality.Parties.Add(party);
                }
            }
        }
        catch
        {
            Console.WriteLine($"Load parties failed for municipality: '{municipality.Name}', URL: {municipality.Url}");
            throw;
        }
    }

    private List<MunicipalityModel> GetMunicipalities(List<CountyModel> counties)
    {
        var municipalities = new List<MunicipalityModel>();
        foreach (var county in counties)
        {
            Console.WriteLine($"Processing county: '{county.Name}'");

            var document = LoadHtmlDocument(county.Url);

            var nodeList = document.DocumentNode.SelectNodes("//td[@class='cislo'][@headers='sa1 sb1']");
            foreach (var node in nodeList)
            {
                var parent = node.ParentNode;
                var urlPath = node.FirstChild.Attributes["href"].Value;
                var nameNode = parent.SelectSingleNode("td[@headers='sa1 sb2']");
                var municipality = new MunicipalityModel()
                {
                    Url = BASE_URL + HttpUtility.HtmlDecode(urlPath),
                    Name = nameNode.InnerText,
                    County = county.Name,
                    Parties = [],
                };
                municipalities.Add(municipality);
                Console.WriteLine($"Identified municipality '{municipality.Name}'");

                LoadParties(municipality);
                Console.WriteLine($"Loaded parties for '{municipality.Name}'");

                if (_isDebugMode) break;
            }
        }
        return municipalities;
    }

    private List<CountyModel> GetCounties()
    {
        var document = LoadHtmlDocument(BASE_URL + "kv12?xjazyk=CZ&xid=0");

        var counties = new List<CountyModel>();
        for (int index = 2; index < 16; index++)
        {
            var nodeList = document.DocumentNode.SelectNodes($"//td[@headers='t{index}sa1 t{index}sb2']");
            foreach (var node in nodeList)
            {
                var urlNode = node.ParentNode.SelectSingleNode($"td[@headers='t{index}sa3']");
                var urlPath = urlNode.FirstChild.Attributes["href"].Value;
                var county = new CountyModel()
                {
                    Name = node.InnerText,
                    Url = BASE_URL + HttpUtility.HtmlDecode(urlPath),
                };
                counties.Add(county);
                Console.WriteLine($"Identified county '{county.Name}'");

                if (_isDebugMode) break;
            }
        }
        return counties;
    }

    private List<OmegaModel> FindOmegaCandidates(List<MunicipalityModel> municipalities)
    {
        var omegaData = new List<OmegaModel>();
        foreach (var municipality in municipalities)
        {
            Console.WriteLine($"Processing municipality: '{municipality.Name}'");
            if (municipality.Parties.Count < 5)
            {
                // Skip municipalities with less than 5 parties.
                continue;
            }
            int omegaTotal = 0;
            foreach (var party in municipality.Parties)
            {
                var document = LoadHtmlDocument(party.Url);

                int? omegaVotes = null;
                var nodeList = document.DocumentNode.SelectNodes("//td[@class='cislo'][@headers='sa2 sb3']");
                foreach (var node in nodeList)
                {
                    var votes = GetIntValue(node);
                    if (omegaVotes == null || omegaVotes.Value > votes)
                    {
                        omegaVotes = votes;
                    }
                }
                if (omegaVotes != null)
                {
                    omegaTotal += omegaVotes.Value;
                }
            }

            var omega = new OmegaModel()
            {
                Municipality = municipality.Name,
                County = municipality.County,
                UsedEnvelopes = municipality.UsedEnvelopes,
                ValidVotes = municipality.ValidVotes,
                OmegaVotes = omegaTotal,
                Url = municipality.Url,
            };
            omegaData.Add(omega);
        }
        return omegaData;
    }

    private void WriteToExcel(List<OmegaModel> omegaData)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("OmegaVotes");
        var rowIndex = 0;
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue("Obec");
        row.CreateCell(1).SetCellValue("Okres");
        row.CreateCell(2).SetCellValue("Odevzdané obálky");
        row.CreateCell(3).SetCellValue("Platné hlasy");
        row.CreateCell(4).SetCellValue("Součet hlasů");
        row.CreateCell(5).SetCellValue("Odkaz");
        rowIndex++;

        foreach (var dataRecord in omegaData)
        {
            var municipalityRow = sheet.CreateRow(rowIndex);
            municipalityRow.CreateCell(0).SetCellValue(dataRecord.Municipality);
            municipalityRow.CreateCell(1).SetCellValue(dataRecord.County);
            municipalityRow.CreateCell(2).SetCellValue(dataRecord.UsedEnvelopes);
            municipalityRow.CreateCell(3).SetCellValue(dataRecord.ValidVotes);
            municipalityRow.CreateCell(4).SetCellValue(dataRecord.OmegaVotes);
            municipalityRow.CreateCell(5).SetCellValue(dataRecord.Url);
            rowIndex++;
        }

        var fileName = $"OmegaVotes_{DateTime.Now:yyyyMMddhhmmss}.xls";
        using (var fileData = new FileStream(fileName, FileMode.Create))
        {
            workbook.Write(fileData);
        }
        Console.WriteLine($"XLS file '{fileName}' written.");
    }

    public void ScrapeMunicipality()
    {
        try
        {
            var counties = GetCounties();

            var municipalities = GetMunicipalities(counties);

            var omegaData = FindOmegaCandidates(municipalities);

            WriteToExcel(omegaData);
        }
        catch (Exception exception)
        {
            Console.WriteLine(exception.ToString());
        }
    }
}
