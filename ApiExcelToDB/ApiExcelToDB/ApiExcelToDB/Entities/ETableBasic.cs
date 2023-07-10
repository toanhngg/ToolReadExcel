using System;

namespace ApiExcelToDB.Entities
{
    public class ETableBasic
    {
        public string StockCode { get; set; }
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public float PriorDayClose { get; set; }
        public float wRecordHigh { get; set; }
        public float wRecordLow { get; set; }
        public float AverageOutstandingShares { get; set; }
        public float PrimaryEPS { get; set; }
        public string Notes { get; set; }
        public float AdjustedEPS { get; set; }
        public float PE { get; set; }
        public float Dividend { get; set; }
        public float DividendMarketPrice { get; set; }
        public float ReturnOnTotalAssets { get; set; }
        public float ReturnOnEquity { get; set; }
        public float ListedShares { get; set; }
        public float OutstandingShares { get; set; }
        public float ChangeOutstandingShares { get; set; }
        public float AdjustedOutstandingShares { get; set; }
        public float TurnoverRatio { get; set; }
        public float MtkCap { get; set; }
    }
}