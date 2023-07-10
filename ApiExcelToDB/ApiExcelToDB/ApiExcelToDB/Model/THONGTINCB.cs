using System;
namespace ApiExcelToDB.Model
{
    public class THONGTINCB
    {
        //STT,Symbol,PriceCloseAverage,KLCPNY,KLCPLH,EPS,ROE,ROA,GTTT
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double PriceCloseAverage { get; set; }
        public double KLCPNY { get; set; }
        public double KLCPLH { get; set; }
        public double EPS { get; set; }
        public double ROE { get; set; }
        public double ROA { get; set; }
        public double GTTT { get; set; }

        public DateTime Trangding_Date { get; set; }
      
    }
}
