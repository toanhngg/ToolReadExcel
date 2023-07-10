using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPNYGTMAX_HNX
    {
        //Symbol,ClosePrice,KLGD,GTNY,Trangding_Date
        public string Symbol { get; set; }
        public double ClosePrice { get; set; }
        public double KLGD { get; set; }
        public double GTNY { get; set; }


        public DateTime Trangding_Date { get; set; }
    }
}
