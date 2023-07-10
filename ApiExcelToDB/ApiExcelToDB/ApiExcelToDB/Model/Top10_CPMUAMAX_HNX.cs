using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPMUAMAX_HNX
    {
        //Symbol,KLGD,GTMUA,KLNG,Trangding_Date
        public string Symbol { get; set; }
        public double KLGD { get; set; }
        public double GTMUA { get; set; }
        public double KLNG { get; set; }


        public DateTime Trangding_Date { get; set; }
    }
}
