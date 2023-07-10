using System;
namespace ApiExcelToDB.Model
{
    public class KQGDCHITIET2011
    {
     
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double BasicPrice { get; set; }
        public double OpenPrice { get; set; }
       
        public double ClosePrice { get; set; }
        public double HighPrice { get; set; }
        public double LowPrice { get; set; }
        public double AveragePrice { get; set; }
      
        
        public double NetChange { get; set; }

        public double Volume_BG { get; set; }
        public double Value_BG { get; set; }
        public double AveragePrice_TT { get; set; }
        public double Volume_TT { get; set; }
        public double Value_TT { get; set; }
        public double Volume_TC { get; set; }
        public double Value_TC { get; set; }
        public double GiaTriTT { get; set; }
      
        public DateTime Trangding_Date { get; set; }
    }
}
