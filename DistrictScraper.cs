using HtmlAgilityPack;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Volby.Model;

namespace Volby
{
    public class DistrictScraper
    {
        private List<string> PartyListEuro = new List<string>() { "Občanská demokratická strana", "Česká pirátská strana", "Koalice STAN, TOP 09", "ANO 2011", "Svob.a př.dem.-T.Okamura (SPD)", "Křes».demokr.unie-Čs.str.lid.", "Komunistická str.Čech a Moravy" };
        private List<string> PartyListMunicipality = new List<string>() { "Občanská demokratická strana", "Zelení a Piráti pro 13", "LIDOVCI (KDU-ČSL) A STAN", "ANO 2011", "TOP 09", "Svob.a př.dem.-T.Okamura (SPD)", "Komunistická str.Čech a Moravy", "Svobodní z Prahy 13", "Společ.proti výst.v Prok.údolí", "Česká str.sociálně demokrat." };

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

        private void WriteToExcel(List<District> districtList, List<string> partyList)
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("Praha 13");
            var rowIndex = 0;
            var row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("Okrsek");
            for (int index = 0; index < partyList.Count; index++)
            {
                row.CreateCell(1 + index).SetCellValue(partyList[index]);
            }
            row.CreateCell(1 + partyList.Count).SetCellValue("Voliči");
            row.CreateCell(2 + partyList.Count).SetCellValue("Hlasy");
            rowIndex++;

            foreach (var district in districtList)
            {
                var districtRow = sheet.CreateRow(rowIndex);
                districtRow.CreateCell(0).SetCellValue(district.Code);
                for (int index = 0; index < partyList.Count; index++)
                {
                    districtRow.CreateCell(1 + index).SetCellValue(district.PartyList[index].Result);
                }
                districtRow.CreateCell(1 + partyList.Count).SetCellValue(district.TotalVoters);
                districtRow.CreateCell(2 + partyList.Count).SetCellValue(district.Voted);
                rowIndex++;
            }

            using (var fileData = new FileStream("Vysledky.xls", FileMode.Create))
            {
                workbook.Write(fileData);
            }
            Console.WriteLine("XLS file written.");
        }

        public void ScrapeEuro()
        {
            var districtList = new List<District>();
            try
            {
                using (var webClient = new WebClient()) // WebClient class inherits IDisposable
                {
                    for (int index = 0; index < 57; index++)
                    {
                        var district = new District()
                        {
                            Code = (13000 + (index + 1)).ToString(),
                            PartyList = new List<PartyResult>(),
                        };

                        var url = $"https://volby.cz/pls/ep2019/ep1311?xjazyk=CZ&xobec=539694&xokrsek={district.Code}&xvyber=1100";
                        string htmlCode = webClient.DownloadString(url);
                        var document = new HtmlDocument();
                        document.LoadHtml(htmlCode);
                        district.TotalVoters = GetCellIntValue(document, "//td[@class='cislo'][@headers='sa2']");
                        district.Voted = GetCellIntValue(document, "//td[@class='cislo'][@headers='sa6']");
                        foreach (var party in PartyListEuro)
                        {
                            var row = document.DocumentNode.SelectSingleNode($"//td[text()='{party}']").ParentNode;
                            foreach (var childNode in row.ChildNodes)
                            {
                                if (childNode.NodeType == HtmlNodeType.Element && (childNode.Attributes["headers"].Value == "t1sa2 t1sb4" || childNode.Attributes["headers"].Value == "t2sa2 t2sb4"))
                                {
                                    district.PartyList.Add(new PartyResult()
                                    {
                                        PartyName = party,
                                        Result = GetDoubleValue(childNode),
                                    });
                                }
                            }
                        }
                        districtList.Add(district);
                    }
                }
                WriteToExcel(districtList, PartyListEuro);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
            foreach (var district in districtList)
            {
                Console.WriteLine($"{district.Code} - {district.Voted} / {district.TotalVoters}");
                foreach (var partyResult in district.PartyList)
                {
                    Console.WriteLine($"{partyResult.PartyName} - {partyResult.Result}");
                }
            }
        }

        public void ScrapeMunicipality()
        {
            var districtList = new List<District>();
            try
            {
                using (var webClient = new WebClient()) // WebClient class inherits IDisposable
                {
                    for (int index = 0; index < 57; index++)
                    {
                        var district = new District()
                        {
                            Code = (13000 + (index + 1)).ToString(),
                            PartyList = new List<PartyResult>(),
                        };

                        var url = $"https://volby.cz/pls/kv2018/kv1111?xjazyk=CZ&xid=0&xdz=5&xnumnuts=1100&xobec=539694&xokrsek={district.Code}&xstat=0&xvyber=1";
                        string htmlCode = webClient.DownloadString(url);
                        var document = new HtmlDocument();
                        document.LoadHtml(htmlCode);
                        district.TotalVoters = GetCellIntValue(document, "//td[@class='cislo'][@headers='sa4']");
                        district.Voted = GetCellIntValue(document, "//td[@class='cislo'][@headers='sa7']");
                        foreach (var party in PartyListMunicipality)
                        {
                            var row = document.DocumentNode.SelectSingleNode($"//td[text()='{party}']").ParentNode;
                            foreach (var childNode in row.ChildNodes)
                            {
                                if (childNode.NodeType == HtmlNodeType.Element && childNode.Attributes["headers"].Value == "t2sa3 t2sb4")
                                {
                                    district.PartyList.Add(new PartyResult()
                                    {
                                        PartyName = party,
                                        Result = GetDoubleValue(childNode),
                                    });
                                    break;
                                }
                            }
                        }
                        districtList.Add(district);
                    }
                }
                WriteToExcel(districtList, PartyListMunicipality);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
        }

    }
}
