using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPGIAMGIA
    {
        //Symbol,MucGIAM,TyLeGiam,KLGD,Trangding_Date
        public string Symbol { get; set; }
        public double MucGIAM { get; set; }
        public double TyLeGiam { get; set; }

        public double KLGD { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
