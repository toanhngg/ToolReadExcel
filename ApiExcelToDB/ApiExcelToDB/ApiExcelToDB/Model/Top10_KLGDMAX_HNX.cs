using System;
namespace ApiExcelToDB.Model
{
    public class Top10_KLGDMAX_HNX
    {
        //Symbol,ClosePrice,KLGD,TyTrong,Trangding_Date
        public string Symbol { get; set; }
        public double ClosePrice { get; set; }
        public double KLGD { get; set; }
        public double TyTrong { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
