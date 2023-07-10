using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPTPRICE
    {
        //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date
        public string Symbol { get; set; }
        public double MucTang { get; set; }
        public double TyLeTang { get; set; }

        public double KLGD { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
