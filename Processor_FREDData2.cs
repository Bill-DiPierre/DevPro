using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Aspose.Cells;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;


namespace BG.Processors
{
    class Processor_FREDData2 : ProcessorBase
    {
        private string _dataFolder = "c:\\saver_temp\\FRED\\";

        public Processor_FREDData2() : base("Fred")
        {

        }
        public override ParsingReturnStructure RunScraper()
        {
            var fileIds = new List<string>() {"CPIAUCSL", "CPIUFDNS", "DGS10", "DJIA", "EXHOSLUSM495S", "FEDFUNDS", "MORTGAGE30US", "MSPNHSUS", "PNFI", "PRFI", "VIXCLS" };
            var i = 0;
            var multiPRS = new ParsingReturnStructure();
            foreach (var fileId in fileIds)
            {
                var fileName = DownloadFile(fileId);
                var pt = CreateParsedTable(fileName);
                var prs = ReadData(pt, fileId);
                //DumpParsingReturnStructure(prs, fileId);
                //new PRSFactory().Save(prs, new System.IO.FileStream(@"c:\saver_temp\prs\Fred" + i++.ToString() + ".prs", System.IO.FileMode.Create));
                multiPRS.Append(prs);
            }
            multiPRS.NeumKey = _nk;
            return multiPRS;
        }

        private string DownloadFile(string fileId)
        {
            try
            {
                var uri = new Uri($"https://fred.stlouisfed.org/graph/fredgraph.xls");
                var localFile = $"{_dataFolder}{DateTime.Now.ToString("yyyyMMdd - HHmmss.")}{fileId}.xls";
                using (WebClient webClient = new WebClient())
                {
                    webClient.QueryString.Add("id", fileId);
                    webClient.DownloadFile(uri, localFile);
                }
                return localFile;
            }
            catch(Exception ex)
            {
                Debug.Assert(false, ex.Message);
                return null;
            }
        }

        private ParsedTable CreateParsedTable(string ExcelFile)
        {
            var fstream = new FileStream(ExcelFile, FileMode.Open);
            var workbook = new Aspose.Cells.Workbook(fstream);
            var sheet = workbook.Worksheets[0];
            var pt = new ParsedTable(sheet);
            fstream.Close();
            return pt;
        }
        private PeriodTypes2017 GetDailyFrequency (string dateString)
        {
            //TODO: maybe this should use last date to determine day, not first? (would capture most recent data, but freq would change every day)

            DateTime date = new DateTime();
            if (!DateTime.TryParse(dateString, out date))
            {
                Debug.Assert(false, "unexpected error parsing date " + dateString);
                return PeriodTypes2017.NotSet;  //TODO log errors
            }
            switch (date.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    return PeriodTypes2017.WeeklySun;
                case DayOfWeek.Monday:
                    return PeriodTypes2017.WeeklyMon;
                case DayOfWeek.Tuesday:
                    return PeriodTypes2017.WeeklyTues;
                case DayOfWeek.Wednesday:
                    return PeriodTypes2017.WeeklyWeds;
                case DayOfWeek.Thursday:
                    return PeriodTypes2017.WeeklyThurs;
                case DayOfWeek.Friday:
                    return PeriodTypes2017.WeeklyFri;
                case DayOfWeek.Saturday:
                    return PeriodTypes2017.WeeklySat;
                default:
                    Debug.Assert(false, "unexpected day of week parsing date " + dateString);
                    return PeriodTypes2017.NotSet;
            }
        }
        private ParsingReturnStructure ReadData(ParsedTable pt, string fileId)
        {
            int firstDateRow = 0;
            PeriodTypes2017 frequency = PeriodTypes2017.Monthly; 
            string seriesName = "";
            var increment = 1;
            //get high level info from the first few rows
            for (int row = 0; row <= pt.Vals.GetUpperBound(0); row++)
            {
                var cell = pt.Vals[row, 0].ToString();

                if (cell.Contains(fileId))
                {
                    seriesName = pt.Vals[row, 1].ToString().Replace(",", " ");  //TODO - probably need to shorten this
                }
                if (cell.Contains("Frequency"))
                {
                    if (cell.Contains("Monthly"))
                    {
                        frequency = PeriodTypes2017.Monthly;
                    }
                    else if (cell.Contains("Quarterly"))
                    {
                        frequency = PeriodTypes2017.Quarterly;
                    }
                    else if (cell.Contains("Weekly"))
                    {
                        frequency = GetDailyFrequency(pt.Vals[row + 2, 0]);  //actual dates start two rows down  ;

                    }
                    else if (cell.Contains("Daily"))
                    {
                        frequency = GetDailyFrequency(pt.Vals[row+2,0]);  //actual dates start two rows down  
                        increment = 5;
                    }
                    else
                    {
                        //TODO - handle more cases
                        Debug.Assert(false, "unknown frequency");
                    }
                }
                DateTime dateResult;
                if (DateTime.TryParse(pt.Vals[row, 0], out dateResult))
                {
                    firstDateRow = row;
                    break;
                }
            }
            var neumonic = SeriesToNeumonic(seriesName);
            seriesName = neumonic;
            Guid seriesId = _nk.GetValue(neumonic);
            _nk.AddSeriesName(seriesId, seriesName);
            var dataPointList = new DataPointList();
            for (int row = firstDateRow; row <= pt.Vals.GetUpperBound(0); row+=increment)
            {
                var dp = new DataPoint2017(DateTime.Parse(pt.Vals[row, 0]), frequency, decimal.Parse(pt.Vals[row, 1]));
                dp.Neum = neumonic;
                dp.ParentSeriesId = seriesId;
                dataPointList.AddPoint(dp);
                
            }
            var prs = new ParsingReturnStructure();
            prs.DataPoints = dataPointList;


            BGTableInfo bgTableInfo = new BGTableInfo();  ///
            bgTableInfo.TableName = seriesName;
            prs.TableInfos.Add(bgTableInfo);
            BGTableLineInformation tableLineInfo = new BGTableLineInformation();
            tableLineInfo.linelabel = seriesName;
            tableLineInfo.tablelineindents = 0;
            tableLineInfo.tablelinenum = 1;
            tableLineInfo.objectID = seriesId;
            bgTableInfo.Add(tableLineInfo);

            return prs;
        }

        private string SeriesToNeumonic(string seriesName)
        {
            //shorten a series name into an understandable neumonic
            //TODO - make this better
            return seriesName.Replace(" ", "");
        }

        private bool DumpParsingReturnStructure(ParsingReturnStructure prs, string fileId)
        {
            //dump the prs just so we can check it

            const string Delimiter = ",";
            var stream = new StreamWriter($"{_dataFolder}dump.{fileId}.csv");

            stream.WriteLine("Date,Neum,Parent,Value");
            for (int i = 0; i < prs.DataPoints.Values.Count; i++)
            {
                var dp = prs.DataPoints.Values[i];
                stream.Write(dp.EndDate);
                stream.Write(Delimiter);
                stream.Write(dp.Neum);
                stream.Write(Delimiter);
                stream.Write(dp.ParentSeriesId);
                stream.Write(Delimiter);
                stream.Write(dp.Val);
                stream.Write("\r\n");
            }

            stream.Flush();
            stream.Close();

            return true;
        }
    }
}
