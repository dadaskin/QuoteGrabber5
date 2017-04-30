
namespace QuoteGrabber5
{
    public class StockInformation
    {
        #region Properties

        public string Symbol { get; private set; }
        public bool ParseAsFund { get; private set; }
        public string RowStr { get; set; }
        public string PricePerShareStr { get; set; }
        public string AnnualDividend { get; set; }
        public string SheetName { get; set; }

        #endregion

        public StockInformation(string symbol, bool parseAsFund, string sheetName, string row)
        {
            Symbol = symbol;
            ParseAsFund = parseAsFund;
            SheetName = sheetName;
            RowStr = row;
        }
    }
}
