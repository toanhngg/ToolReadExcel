using System;
namespace ApiExcelToDB.Model
{
    public class Top10CK_TANGGIA2010
    {
      //Symbol,AvePrice,MucTang,TyLeTang,Trangding_Date
        public string Symbol { get; set; }
      
        public double AvePrice { get; set; }
        public double MucTang { get; set; }

        public double TyLeTang { get; set; }
  
        public DateTime Trangding_Date { get; set; }
    }
}
