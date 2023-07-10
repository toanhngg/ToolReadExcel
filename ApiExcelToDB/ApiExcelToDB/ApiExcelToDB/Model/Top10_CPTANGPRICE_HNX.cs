using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPTANGPRICE_HNX
    {
        //Symbol,ClosePrice,MucTang,TyLeTang,KLGD,Trangding_Date
        public string Symbol { get; set; }
        public double ClosePrice { get; set; }
        public double MucTang { get; set; }
        public double TyLeTang { get; set; }

        public double KLGD { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
