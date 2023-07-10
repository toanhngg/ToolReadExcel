using System;

namespace ApiExcelToDB.Entities
{
    public class ETableConstituents
    {
        public string StockCode { get; set; }
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public float TodayClose { get; set; }
        public float OutstandingShares { get; set; }
        public float ShareRestrictedOnTransfer { get; set; }
        public float FreeFloat { get; set; }
        public float CapRatio { get; set; }
        public float FreeFloatAdjustedMarketCap { get; set; }
        public float Weight { get; set; }
        public string Type { get; set; }

        public float PriceClose { get; set; }
    }
}
