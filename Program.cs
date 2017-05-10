// *********
// Change Log:
// 15 Aug 2011  v3.0:   Detailed view for mutual funds no longer provides closing price. So I now make two HTML requests: One
//                      detailed view for stocks that provides closing price and div yield, and one basic view that provides the 
//                      closing price.  I'll assume no dividend yield for mutual funds.  Any issue who's symbol ends in "X" is 
//                      considered a Fund for this application (except FAX).  Also JNK in included as a Fund.
//
//  2 May 2012  v3.1:   Yahoo changed the search string from  to "yfs_l84_..."
//
// 13 Jul 2012  v3.2:   Now I have to send each fund request separately.
//                      All funds except JNK have a search string of "yfs_l10_..."
//                      Search string for JNK is "yfs_l84_".
//
// 26 Nov 2012  v4.0:   Port to VSExpress 2012 and new computer.
//
// 15 Apr 2014  v4.0.1: In readSymbolsFromSheet() change ÜTX to KERX for symbols ending in "X" to parse as stocks.
//
// 11 Mar 2015  v4.1.0: Re-throw WebExceptions and try/catch at top level.
//
// 21 Mar 2016  v4.1.1: Parse GLD as a stock, but skip over the yield.  Not doing so messes up the parsing for the following stocks.
//
// 12 Aor 2016  v4.1.2: Add SBUX to exception list for symbols ending in "X" that aren't funds.
//
// 13 Jul 2016  v4.1.3:  - Now we can go back to getting all fund quotes en masse.
//                      - Get fund quotes as tabluar "v1" view (not detailed "dv" view)
//
// 19 Jul 2016  v4.1.4: Add NFLX to exception list for symbols ending in "X" that aren't funds.
//
// 20 Apr 2017  v5.0.0: As of 28 Apr 2017 Yahoo Finance got rid of it's detailed view option.
//                      To continue to get the quotes I need to radically revamp QuoteGrabber to use the 
//                      Yahoo Query Interface (YQI) to get the information in JSON format and then change
//                      the logic in the program to parse JSON and update the spreadsheet.
//
// 09 May 2017  v5.1.0: Change how last symbol row is found.  Was: row of last symbol on Sheet1. Didn't
//                      work if last n rows was different blocks of same symbol.
//

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuoteGrabber5
{
    class Program
    {
        static void Main()
        {
// ReSharper disable ObjectCreationAsStatement
            new Program();
// ReSharper restore ObjectCreationAsStatement
        }

        private readonly IList<StockInformation> _mInvestments = new List<StockInformation>();
        private readonly string _mSsfn;
        private string _mDateStr;
        private int _mLastSymbolRow;

        // Excel Spreadsheet landmarks
        private const string MSymbolColStr = "A";
        private const string MPricePerShareCol = "C";
        private const string MAnnualDividendCol = "N";
        private const int MFirstSymbolRow = 5;
        private const string MSymbolListTerminationString = "Cash";
        private const string MTimeStampCell = "C1";

        Excel.Application _oXl;
        Excel._Workbook _oWb;

        public Program()
        {
            Console.WriteLine("QuoteGrabber v5.0.0 - (c)2017 David R. Adaskin");
            _mSsfn = "C:\\Users\\Dave\\Documents\\ThirdTestPortfolio.xls";

            ReadSymbolsFromSpreadsheet();

            try
            {
                DoWebRequestAndParse();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                Console.WriteLine("Press the any key to terminate...");
                Console.ReadLine();
                return;
            }

            DisplayResults();

            UpdateSpreadsheet();

            Console.WriteLine("Press the any key to terminate...");
            Console.ReadLine();
        }

        private void DisplayResults()
        {
            foreach (var info in _mInvestments)
            {
                Console.WriteLine(info.Symbol + "\t" 
                                  + info.SheetName + "\t"
                                  + info.RowStr + "\t"
                                  + info.PricePerShareStr + "\t"
                                  + info.AnnualDividend);                
            }
        }

        private void UpdateSpreadsheet()
        {
            var nextAvailRowCell = ((char)(3 + 64)) + (_mLastSymbolRow + 7).ToString(CultureInfo.InvariantCulture);
            var totalValueCell = ((char)(4 + 64)) + (_mLastSymbolRow + 4).ToString(CultureInfo.InvariantCulture);
            var avgAgeCell = ((char)(20 + 64)) + (_mLastSymbolRow + 4).ToString(CultureInfo.InvariantCulture);
            var avgReturnEsppCell = ((char)(24 + 64)) + (_mLastSymbolRow + 4).ToString(CultureInfo.InvariantCulture);
            var avgReturnNoEsppCell = ((char)(25 + 64)) + (_mLastSymbolRow + 4).ToString(CultureInfo.InvariantCulture);

            try
            {
                var wSheet1 = (Excel._Worksheet)_oWb.Sheets[1];
                var wSheet2 = (Excel._Worksheet)_oWb.Sheets[2];

                var oWs = wSheet1;

                // Write the date at the top
                var tsRng = oWs.Range[MTimeStampCell, MTimeStampCell];

                double oldDate = 0;
                if (tsRng.Value2 != null)
                    oldDate = (double)tsRng.Value2;
                tsRng.Formula = _mDateStr;
                tsRng.NumberFormat = "dd-mmm-yy";

                var newDate = (double)tsRng.Value2;

                // If we have already stored values for this date don't do it again
                if (Math.Abs(Math.Floor(oldDate) - Math.Floor(newDate)) > .0001)
                {
                    // Write Price values to correct cells
                    //foreach (string s in mInfo.Keys)
                    foreach (StockInformation issue in _mInvestments)
                    {
                        oWs = issue.SheetName == "Sheet1" ? wSheet1 : wSheet2;

                        var ssCell = MPricePerShareCol + issue.RowStr;
                        var ppsRng = oWs.Range[ssCell, ssCell];
                        ppsRng.Formula = issue.PricePerShareStr;

                        ssCell = MAnnualDividendCol + issue.RowStr;
                        var divRng = oWs.Range[ssCell, ssCell];
                        divRng.Formula = issue.AnnualDividend;
                    }

                    Console.WriteLine("Values updated...");

                    oWs = wSheet1;

                    // Read "Next Available Row" value and parse to an int
                    var nextAvailRng = oWs.Range[nextAvailRowCell, nextAvailRowCell];
                    var row = Int32.Parse(nextAvailRng.Value2.ToString());

                    // Go to col 1 of this row and paste the date
                    oWs.Cells[row, 1] = tsRng.Formula;

                    // Get the total value and copy that in the 2nd column of the cum. list
                    var rngSrc = oWs.Range[totalValueCell, totalValueCell];
                    oWs.Cells[row, 2] = rngSrc.Value2;

                    // Copy the percent gain/loss formula from the previous row to this row
                    string srcCell = ((char)(4 + 64)) + (row - 1).ToString();
                    string desCell = ((char)(4 + 64)) + row.ToString();
                    rngSrc = oWs.Range[srcCell, srcCell];
                    var rngDes = oWs.Range[desCell, desCell];
                    rngSrc.Copy(rngDes);

                    // Copy the avg. rtn with ESPP to the 12th column 
                    rngSrc = oWs.Range[avgReturnEsppCell, avgReturnEsppCell];
                    oWs.Cells[row, 12] = rngSrc.Value2;

                    // Copy the avg. rtn w/o ESPP to the 13th column
                    rngSrc = oWs.Range[avgReturnNoEsppCell, avgReturnNoEsppCell];
                    oWs.Cells[row, 13] = rngSrc.Value2;

                    // Copy the avg age to the 14th column
                    rngSrc = oWs.Range[avgAgeCell, avgAgeCell];
                    oWs.Cells[row, 14] = rngSrc.Value2;

                    // Fix the color of the cumulative total
                    srcCell = ((char)(2 + 64)) + row.ToString();
                    rngSrc = oWs.Range[srcCell, srcCell];
                    rngSrc.Font.Bold = false;
                    rngSrc.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

                    Console.WriteLine("Values placed in Cumulative table...");

                    // Increment next available row and place it in the spreadsheet
                    row++;
                    nextAvailRng.Formula = row;
                }
                else
                {
                    Console.WriteLine("oldDate: {0}({1})   newDate: {2}({3})",
                                      oldDate, Math.Floor(oldDate),
                                      newDate, Math.Floor(newDate));
                    Console.WriteLine("Not updating again today.");
                }
            }
            catch (Exception err)
            {
                string eMsg = "Error:  ";
                eMsg += string.Concat(eMsg, err.Message);
                eMsg += string.Concat(eMsg, " Line: ");
                eMsg += string.Concat(eMsg, err.Source);
                Console.WriteLine(eMsg);
            }

        }

        #region Read Spreadsheet methods
        private void ReadSymbolsFromSpreadsheet()
        {
            try
            {
                // Open the spreadsheet
                _oXl = new Excel.Application { Visible = true };
                var z = Missing.Value;
                _oWb = _oXl.Workbooks.Open(_mSsfn, z, z, z, z, z, z, z, z, z, z, z, z, z, z);

                foreach (var obj in _oWb.Sheets)
                {
                    ReadSymbolsFromSheet((Excel._Worksheet)obj);
                }

            //    GetLastSymbolRowOnSheet1();
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        //private void GetLastSymbolRowOnSheet1()
        //{
        //    int prevRow = 0;
        //    foreach (var issue in _mInvestments)
        //    {
        //        if (issue.SheetName == "Sheet2")
        //        {
        //            _mLastSymbolRow = prevRow;
        //            break;
        //        }
        //        prevRow = Int32.Parse(issue.RowStr);
        //    }

        //}


        // TBD: Simplify  don't need parseAs

        private void ReadSymbolsFromSheet(Excel._Worksheet sheet)
        {
            var row = MFirstSymbolRow;

            var symbol = string.Empty;
            while (symbol != MSymbolListTerminationString)
            {
                var cell = MSymbolColStr + row.ToString(CultureInfo.InvariantCulture);
                var symbolRng = sheet.Range[cell, cell];
                symbol = (string)symbolRng.Value2;

                symbol = CleanSymbol(symbol);

                if ((symbol != null) && (symbol != MSymbolListTerminationString))
                {
                    var parseAsFund = (symbol.EndsWith("X") &&
                                        (symbol != "CVX") &&
                                        (symbol != "FAX") &&
                                        (symbol != "SBUX") &&
                                        (symbol != "NFLX") &&
                                        (symbol != "FCX"));

                    var issue = new StockInformation(symbol, parseAsFund, sheet.Name, row.ToString(CultureInfo.InvariantCulture));

                    var alreadyInList = _mInvestments.Any(item => item.Symbol == issue.Symbol);

                    if (!alreadyInList)
                        _mInvestments.Add(issue);
                }
                else if ((symbol != null) && (symbol == MSymbolListTerminationString) && sheet.Name.Contains("Sheet1"))
                {
                    _mLastSymbolRow = row - 1;
                }
                row++;
            }
        }

        // TBD: Simplify?
        // Remove any extra information from symbol string 
        private static string CleanSymbol(string symbol)
        {
            if (symbol == "Scottrade")
                return null;

            if (symbol == "Shareowner Services")
                return null;

            if ((symbol != null) && (symbol != MSymbolListTerminationString))
            {
                // Remove any extra information from symbol string
                var firstUnnecessaryCharacter = symbol.IndexOf(' ');
                if (firstUnnecessaryCharacter > 0)
                    symbol = symbol.Remove(firstUnnecessaryCharacter);
            }
            return symbol;
        }

        #endregion Read Spreadsheet methods

        #region Get and Parse information methods
        
        private void DoWebRequestAndParse()
        {
            var symbolList = GetSymbolList();

            const string urlBeforeSymbols = @"https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20csv%20where%20url%3D'http%3A%2F%2Fdownload.finance.yahoo.com%2Fd%2Fquotes.csv%3Fs%3D";
            const string urlAfterSymbols = @"%26f%3Dsl1d%26e%3D.csv'%20and%20columns%3D'symbol%2Cprice%2Cdividend'&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys";
            var wholeUrl = urlBeforeSymbols + symbolList + urlAfterSymbols;
            Console.WriteLine("URL: " + wholeUrl);

            var webRequest = (HttpWebRequest) WebRequest.Create(wholeUrl);
            Console.WriteLine("WebRequest sent...");
            try
            {
                Console.WriteLine("Awaiting WebResponse...");
                using (var webResponse = webRequest.GetResponse())
                {
                    Console.WriteLine("WebResponse received...");
                    var responseStream = webResponse.GetResponseStream();
                    if (responseStream != null)
                    {
                        var reader = new StreamReader(responseStream);
                        var sourceString = reader.ReadLine();
                        if (string.IsNullOrEmpty(sourceString))
                            throw new Exception("StreamReader.ReadLine() returns empty string.");

                        _mDateStr = GetMarketDate(sourceString);
                        ParseForEachSymbol(sourceString);
                    }
                }
            }
            // try WebRequest.Create()
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("*** Exception! ***");
                Console.WriteLine(ex.Message);
                Thread.Sleep(5000);
                throw new Exception("Re-throw of Exception");
            }

        }

        private void ParseForEachSymbol(string sourceString)
        {
            const string entrySeparator = ",";
            const string lastEntryFlag = "}";
            const string symbolKey = "\"symbol\"";
            const string ppsKey = "\"price\"";
            const string dividendKey = "\"dividend\"";

            var startingPosition = sourceString.IndexOf(symbolKey, StringComparison.Ordinal);
            var source = sourceString.Substring(startingPosition, sourceString.Length - startingPosition);

            foreach (var issue in _mInvestments)
            {
                var beginSymbol = GetValueBeginIdx(source, symbolKey);
                var endSymbol = source.IndexOf(entrySeparator, StringComparison.Ordinal) - 1;  // Account for closing quote
                
                var symbol = source.Substring(beginSymbol, endSymbol - beginSymbol);
                if (string.CompareOrdinal(issue.Symbol, symbol) != 0)
                    throw new Exception("Expecting: " + issue.Symbol + " but found: " + symbol);

                source = source.Substring(endSymbol+2, source.Length - endSymbol-2);

                var beginPps = GetValueBeginIdx(source, ppsKey);
                var endPps = source.IndexOf(entrySeparator, StringComparison.Ordinal) - 1;  //Account for closing quote
                issue.PricePerShareStr = source.Substring(beginPps, endPps- beginPps);

                source = source.Substring(endPps+2, source.Length - endPps-2);

                var beginDiv = GetValueBeginIdx(source, dividendKey);
                var endDiv = source.IndexOf(entrySeparator, StringComparison.Ordinal) - 2;  //Account for closing quote and closing brace
                if (endDiv < 0)
                    endDiv = source.IndexOf(lastEntryFlag, StringComparison.Ordinal) - 1;
                issue.AnnualDividend = source.Substring(beginDiv, endDiv - beginDiv);

                source = source.Substring(endDiv + 4, source.Length - endDiv-4);

                Console.WriteLine("Found info for " + issue.Symbol);
            }
        }

        private int GetValueBeginIdx(string sourceString, string key)
        {
            return sourceString.IndexOf(key, StringComparison.Ordinal) + key.Length + 2;
        }

        private string GetSymbolList()
        {
            var symbolList = _mInvestments.Aggregate(string.Empty, (current, issue) => current + (issue.Symbol + ","));

            // Remove comma after final symbol
            symbolList = symbolList.Remove(symbolList.Length - 1);

            return symbolList;
        }

        /// <summary>
        /// Get the last market date in Pacific Time from the input stream.
        /// Input stream has date in GMT in the format YYYY-MM-DDTHH:mm:ss of the query in JSON.  
        /// The keyword is "created".
        /// If the query is made on a Saturday or Sunday, the date string corresponding to the previous Friday is returned
        /// This method does not account for holidays where the market is closed on a Friday.
        /// </summary>
        /// <param name="sourceString">The input stream</param>
        /// <returns></returns>
        private static string GetMarketDate(string sourceString)
        {
            const string dtKey = "\"created\"";
            const string nextKey = "\"lang\"";

            var beginDate = sourceString.IndexOf(dtKey, StringComparison.Ordinal)
                            + dtKey.Length
                            + 2;
            var endDate = sourceString.IndexOf(nextKey, StringComparison.Ordinal) - 2;

            var dtString = sourceString.Substring(beginDate, endDate - beginDate);

            const int queryYearIdx = 4;
            const int queryMonthIdx = 7;
            const int queryDayIdx = 10;
            const int queryHourIdx = 13;

            var queryYearStr = dtString.Substring(0,queryYearIdx);
            var queryMonthStr = dtString.Substring(queryYearIdx + 1, queryMonthIdx - queryYearIdx -1);
            var queryDayStr = dtString.Substring(queryMonthIdx + 1, queryDayIdx - queryMonthIdx - 1);
            var queryHourStr = dtString.Substring(queryDayIdx + 1, queryHourIdx - queryDayIdx - 1);

            try
            {
                int queryYear = int.Parse(queryYearStr);
                int queryMonth = int.Parse(queryMonthStr);
                int queryDay = int.Parse(queryDayStr);
                int queryHour = int.Parse(queryHourStr);

                var queryDt = new DateTime(queryYear, queryMonth, queryDay, queryHour, 0, 0);

                var tzAdjustment = DateTime.UtcNow - DateTime.Now;
                var queryTimeInZone = queryDt - tzAdjustment;

                if (queryTimeInZone.DayOfWeek == DayOfWeek.Saturday)
                    queryTimeInZone = queryTimeInZone.AddDays(-1);

                if (queryTimeInZone.DayOfWeek == DayOfWeek.Sunday)
                    queryTimeInZone = queryTimeInZone.AddDays(-2);

                return queryTimeInZone.ToShortDateString();
            }
            catch (Exception)
            {
                Console.WriteLine("Error parsing Query time string: " + dtString);
                throw;
            }
        }


        #endregion Get and Parse information methods
    }
}
