using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPGIAMPRICE_HNX
    {
        //Symbol,ClosePrice,MucGiam,TyLeGiam,KLGD,Trangding_Date
        public string Symbol { get; set; }
        public double ClosePrice { get; set; }
        public double MucGiam { get; set; }
        public double TyLeGiam { get; set; }
        public double KLGD { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
