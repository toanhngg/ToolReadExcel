using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPGDMAX_HNX
    {
        //Symbol,ClosePrice,GTGD,TyTrong,Trangding_Date
        public string Symbol { get; set; }
        public double ClosePrice { get; set; }
        public double GTGD { get; set; }
        public double TyTrong { get; set; }


        public DateTime Trangding_Date { get; set; }
    }
}
