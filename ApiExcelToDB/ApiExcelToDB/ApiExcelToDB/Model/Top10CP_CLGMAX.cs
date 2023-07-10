using System;
namespace ApiExcelToDB.Model
{
    public class Top10CP_CLGMAX
    {
      //Symbol,HighPrice,LowPrice,TyLeChenhLech,Trangding_Date
        public string Symbol { get; set; }
      
        public double HighPrice { get; set; }
        public double LowPrice { get; set; }
        public double TyLeChenhLech { get; set; }
      
        public DateTime Trangding_Date { get; set; }
    }
}
