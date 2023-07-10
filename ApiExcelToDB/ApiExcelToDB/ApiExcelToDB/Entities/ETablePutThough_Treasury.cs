using System;

namespace ApiExcelToDB.Entities
{
    public class ETablePutThough_Treasury
    {
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public string StockCode { get; set; }
        public float TradingVolume { get; set; } 
         public float TradingValue { get; set; }
        //public float TotalVolumeEntireMarket { get; set; }
        //public float  TotalVolumeOrderMatching { get; set; }
        //public float TotalVolumePutThough { get; set; }
        //public float TotalValueEntireMarket { get; set; }
        //public float TotalValueOrderMatching { get; set; }
        //public float TotalValuePutThough { get; set; }\
        public float TotalTrading { get; set; }
        public float RateTotalTrading { get; set; }
        public string Type { get; set; }
    
    }
}
