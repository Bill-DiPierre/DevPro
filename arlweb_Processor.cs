using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace BG.Processors
{
    class arlweb_Processor : ProcessorBase
    {
    // scrapping "https://arlweb.msha.gov/stats/centurystats/coalstats.asp"
    // scrapping "https://arlweb.msha.gov/stats/centurystats/mnmstats.asp"
        public arlweb_Processor() : base("arlweb")
        {

        }

        public override ParsingReturnStructure RunScraper()
        {
            ParsingReturnStructure Mprs = new ParsingReturnStructure();
            var prs1 = DownloadTable2("https://arlweb.msha.gov/stats/centurystats/coalstats.asp"); //Table 106
            var prs2 = DownloadTable2("https://arlweb.msha.gov/stats/centurystats/mnmstats.asp");
            Mprs.Append(prs1);
            Mprs.Append(prs2);
            return Mprs;
        }


        private static String CreateNeum(String neum)
        {
            var modifiedNeum = Regex.Replace(neum, @"\n+", "");

            if (String.IsNullOrEmpty(modifiedNeum)) throw new ArgumentNullException("neum");

            var i = 0;
            var first = 0;
            var newNuem = "";
            while (i < modifiedNeum.Length)
            {
                foreach (var letter in modifiedNeum)
                {

                    if (letter == ' ' || letter == modifiedNeum[modifiedNeum.Length - 1])
                    {
                        char firstLetter = modifiedNeum[first];
                        if (Char.IsUpper(firstLetter))
                        {
                            newNuem = newNuem + firstLetter;
                        }
                        if (Char.IsLower(firstLetter))
                        {
                            newNuem = newNuem + Char.ToUpper(firstLetter);
                        }
                        first = i + 1;
                    }
                    i++;
                }
            }
            return Convert.ToString(newNuem);
        }
        private static DateTime FormattedDate(String unformattedDate)
        {
            int year = 0;
            if (unformattedDate.Contains("-"))
            {
                var splitDate = unformattedDate.Split('-');
                year = Convert.ToInt32(splitDate[0]);
            }
            else
            {
                //Get rid of non digits
                year = Convert.ToInt32(Regex.Replace(unformattedDate, "[^0-9]", ""));
                String yearString = year.ToString();
                if (yearString.Length > 4)
                {
                    year = Convert.ToInt32(yearString.Substring(0, 4));
                }

            }

            // Convert to date object
            DateTime dateObject = new DateTime(year, 12, 31);
            // Subtract 1 year and add one day to get last day of December
            return dateObject;

        }
        private ParsingReturnStructure DownloadTable2(String http)
        {
            WebClient webClient = new WebClient();
            string page = webClient.DownloadString(http);
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(page);
            //Console.WriteLine("ABCDEFG".IndexOf("H"));
            string tablescript = page.Substring(page.IndexOf("BORDERDARK") + "BORDERDARK".Length, page.IndexOf("</TABLE>") - page.IndexOf("BORDERDARK") - "BORDERDARK".Length).Split(new string[] { "<tr>" }, StringSplitOptions.None)[1];
            //Console.WriteLine(tablescript);
            var Items = tablescript.Split(new string[] { "<FONT STYLE=\"FONT-SIZE:.90em;\">" }, StringSplitOptions.None).ToList();
            Items.RemoveAt(0);
            //Console.WriteLine(Items[0] + "\n-------------------------");
            //Console.WriteLine(Items[15] + "\n-------------------------");
            List<List<String>> Pseudo_Table = new List<List<String>>();
            for (int i = 0; i < Items.Count; i++)
            {
                Pseudo_Table.Add(Regex.Replace(Items[i], "(<.*>)", "").Replace("\"", "").Trim().Split(new string[] { "\n" }, StringSplitOptions.None).ToList());
            }
            Console.WriteLine(Pseudo_Table[22][0] == "");
            for (int i = 0; i < Pseudo_Table.Count; i++)
            {
                if (Pseudo_Table[i][0] == "")
                {
                    Pseudo_Table.RemoveAt(i);
                }
            }


            List<String> Year = new List<String>();
            List<String> Item1 = new List<String>();
            List<String> Item2 = new List<String>();
            int N = Pseudo_Table.Count / 2;

            for (int i = 0; i < N / 3; i++)
            {
                Year.AddRange(Pseudo_Table[N + i * 3]);
                Item1.AddRange(Pseudo_Table[N + i * 3 + 1]);
                Item2.AddRange(Pseudo_Table[N + i * 3 + 2]);
            }
            List<List<String>> table = new List<List<String>>();
            table.Add(new List<String> { Pseudo_Table[0][0], Pseudo_Table[1][0], Pseudo_Table[2][0] });
            for (int i = 0; i < Year.Count; i++)
            {
                table.Add(new List<String> { Year[i].Trim(), Item1[i].Trim(), Item2[i].Trim() });
            }
            String Trpage = page.Substring(page.IndexOf("pageheader\">") + "pageheader\">".Length, 1000);
            String Header = Trpage.Substring(0, Trpage.IndexOf("</div>")).Trim();
            //Console.WriteLine(Header);
            table[0][0] = Header;
            //Console.WriteLine(FormattedDate(table[1][0]));
            //Console.WriteLine(Convert.ToDecimal(""));
            return ProcessData2(table);
        }
        private ParsingReturnStructure ProcessData2(List<List<String>> table)
        {
            NeumKey _nk = new NeumKey();
            DataPointList dataPointList = new DataPointList();
            List<String> footnotes = new List<string>();
            List<String> columnHeaders = new List<string>();
            List<DateTime> dates = new List<DateTime>();
            List<String> values = new List<String>();
            List<Tuple<String, String, DateTime, String, String>> dataPoint = new List<Tuple<string, string, DateTime, string, string>>();
            List<Tuple<String, String>> nuemWithValue = new List<Tuple<string, string>>();
            List<String> neums = new List<String>();


            var rawNeum = table[0][0].ToString();
            var formattedNeum = CreateNeum(rawNeum);
            //var usedRange = ws.UsedRange;
            foreach (List<String> Row in table)
            {
                int Rownum = table.IndexOf(Row);
                foreach (String Item in Row)
                {
                    int Colnum = Row.IndexOf(Item);
                    //Console.WriteLine(Item);
                    if ((Rownum == 0) && (Colnum >= 1))
                    {
                        columnHeaders.Add(Item.ToString());
                    }
                    else if ((Rownum > 0) && (Colnum == 0))
                    {
                        dates.Add(FormattedDate(Item.ToString()));
                    }
                    else if ((Rownum > 0) && (Colnum > 0))
                    {
                        try
                        {
                            Convert.ToDecimal(Item);
                            values.Add(Item);
                        }
                        catch
                        {
                            try
                            {
                                Convert.ToDecimal(Item.Split(')')[1].ToString().Trim());
                                //Console.WriteLine(Item.Split(')')[1].ToString().Trim());
                                values.Add(Item.Split(')')[1].ToString().Trim());
                            }
                            catch
                            {
                                values.Add("-999,999");
                            }

                        }
                        ;
                    }
                }
            }
            foreach (var col in columnHeaders)
            {
                nuemWithValue.Add(new Tuple<string, string>((formattedNeum + "_" + col), (rawNeum + "_" + col)));
            }
            var i = 0;
            foreach (var date in dates)
            //foreach (Tuple<string, string> nuemWithVal in nuemWithValue)
            {

                //foreach(var date in dates)
                foreach (Tuple<string, string> nuemWithVal in nuemWithValue)
                {

                    String[] neumAndCatsubstrings = nuemWithVal.Item1.Split('_');
                    var dp = new DataPoint2017(date, PeriodTypes2017.Annual, Convert.ToDecimal(values[i]));
                    dp.Neum = nuemWithVal.Item1;
                    Guid seriesId = _nk.GetValue(dp.Neum);
                    String seriesName = nuemWithVal.Item1;
                    _nk.AddSeriesName(seriesId, seriesName);
                    dp.ParentSeriesId = seriesId;
                    //dataPoint.Add(new Tuple<string, string, DateTime, string, string>(nuemWithVal.Item1, neumAndCatsubstrings[1], date, "Fiscal Year", values[i]));
                    dataPointList.AddPoint(dp);
                    i++;
                }

            }
            ParsingReturnStructure toReturn = new ParsingReturnStructure();
            BGTableInfo tableHere = new BGTableInfo();
            tableHere.TableName = rawNeum;
            int tableLineNum = 0;
            toReturn.DataPoints.AddList(dataPointList);
            for (int rowOn = 1; rowOn < columnHeaders.Count; rowOn++)
            {
                //      int colOn = 1;

                if (columnHeaders[rowOn] != null && columnHeaders[rowOn].Length > 0)  //sometimes there are blank cells that don't need to be processed
                {
                    tableLineNum++;



                    String seriesName = rawNeum + "_" + columnHeaders[rowOn];


                    String neum = "arlweb~" + seriesName;
                    neum = _nk.PrettifyNeum(neum);
                    Guid seriesID = _nk.GetValue(neum);

                    BGTableLineInformation tableLineInfo = new BGTableLineInformation();
                    tableLineInfo.linelabel = columnHeaders[rowOn];
                    tableLineInfo.tablelineindents = 0;
                    tableLineInfo.tablelinenum = tableLineNum;
                    tableLineInfo.objectID = seriesID;
                    tableHere.Add(tableLineInfo);
                }
            }
            toReturn.NeumKey = _nk;
            toReturn.TableInfos.Add(tableHere);
            return toReturn;
        }
    }
}
