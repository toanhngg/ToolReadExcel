using System;
namespace ApiExcelToDB.Model
{
    public class KLGD_TOP2011_MR
    {
      
        public string Symbol { get; set; }
      
        public double AvePrice { get; set; }
        public double KL { get; set; }
        public double GT { get; set; }
        public double TangGiam { get; set; }
        public double KLGD_NgayTruoc { get; set; }
        //Symbol,AvePrice,KL,GT,TangGiam,KLGD_NgayTruoc,Trangding_Date
        public DateTime Trangding_Date { get; set; }
    }
}
