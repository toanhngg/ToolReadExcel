using System;
namespace ApiExcelToDB.Model
{
    public class UPCoM_KQGD_Phien_2011
    {
       
       
    
        public string GiaoDich { get; set; }
        public string Symbol { get; set; }
        public double BasicPrice { get; set; }
        public double CellingPrice { get; set; }
        public double FloorPrice { get; set; }

        public double HighestPrice { get; set; }
        public double LowestPrice { get; set; }

        public double OpenPrice { get; set; }
        public double ClosePrice { get; set; }
   
        public double Gia_BQ { get; set; }
        public double KLGD_KL { get; set; }
        public double GTGD_KL { get; set; }
        public double HighestPrice_TT { get; set; }
        public double LowestPrice_TT { get; set; }
        public double KLGD_TT { get; set; }
        public double GTGD_TT { get; set; }
        public double KLGD_TC { get; set; }
        public double GTGD_TC { get; set; }
        public double Muc_VHTT { get; set; }

        public double KL_DKGD { get; set; }
        public double KLCPLH { get; set; }
        public double KLMUA { get; set; }
        public double GTMUA { get; set; }

        public double KLBAN { get; set; }
        public double GTBAN { get; set; }
        public double TongKLDPNG { get; set; }
        public double KLCDPNG { get; set; }
       
        public DateTime Trangding_Date { get; set; }
    }
}
