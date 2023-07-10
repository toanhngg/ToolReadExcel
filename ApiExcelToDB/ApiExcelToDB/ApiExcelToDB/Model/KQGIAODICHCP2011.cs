using System;
namespace ApiExcelToDB.Model
{
    public class KQGIAODICHCP2011
    {
        //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
        //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double SLCP_DKGD { get; set; }
        public double SLCP_LH { get; set; }
        public double Co_Tuc_2010 { get; set; }
        public double PE { get; set; }
        public double EPS2010 { get; set; }
        public double KLGD_10PHIEN { get; set; }
        public double ROE { get; set; }

        public double ROA { get; set; }
        public double BasicPrice_KT { get; set; }
        public double CeilingPrice_KT { get; set; }
        public double FloorPrice_KT { get; set; }
        public double BinhQuan { get; set; }
        public double Tong { get; set; }
        public double Co_Tuc_2009 { get; set; }
     
        public DateTime Trangding_Date { get; set; }
    }
}
