//using System.Security.Principal;

using System;
using System.Data;

namespace ApiExcelToDB.Entities
{
    public class ETable 
    { 
        public string StockCode { get; set; }
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public float PriorDayClose { get; set; }
        public float SessionClose { get; set; }
        public float Change { get; set; }
        public float TradingVolume { get; set; }
        public float TradingValue { get; set; }
        public float SessionHigh { get; set; }
        public float SessionAverage { get; set; }
        public float SessionLow { get; set; }
        public float TodayClose { get; set; }
        public float Totalvolume { get; set; }
        public float Totalvalue { get; set; }
        public float ListedShares { get; set; }
        public float OutstandingShares { get; set; }
        public float AdjustedOutstandingShares { get; set; }
        public float Marketcap { get; set; }
        // public string Type { get; set; }
  
        public float UpDown { get; set; }
        public float Low { get; set; }
        public float High { get; set; }
        public float CloseIndexValue { get; set; }
        public float OpenIndexValue { get; set; }
        public string IndexName { get; set; }

    }
}
