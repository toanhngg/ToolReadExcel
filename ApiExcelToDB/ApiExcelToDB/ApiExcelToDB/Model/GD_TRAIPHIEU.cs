using System;
namespace ApiExcelToDB.Model
{
    public class GD_TRAIPHIEU
    {
        //STT,Symbol,KyHanNam,GiaGDDong,LaiSuat,LoiSuat,KLGD,GTGD,Trangding_Date
        public int STT { get; set; }
        public string Symbol { get; set; }
      
        public double KyHanNam { get; set; }
        public double GiaGDDong { get; set; }
        public double LaiSuat { get; set; }

        public double LoiSuat { get; set; }
        public double KLGD { get; set; }

        public double GTGD { get; set; }
      

        public DateTime Trangding_Date { get; set; }
    }
}
