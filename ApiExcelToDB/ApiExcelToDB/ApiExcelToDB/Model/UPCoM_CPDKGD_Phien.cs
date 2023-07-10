using System;
namespace ApiExcelToDB.Model
{
    public class UPCoM_CPDKGD_Phien
    {
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double KLCP_NY { get; set; }
        public double KLCP_LH { get; set; }
        public double Co_Tuc_2010 { get; set; }

        public double PE { get; set; }
        public double EPS2010 { get; set; }
        public double ROE { get; set; }
        public double ROA { get; set; }
        public double BasicPrice_KT { get; set; }
        public double CeilingPrice_KT { get; set; }
        public double FloorPrice_KT { get; set; }
        //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2010,PE,EPS2010,ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
        public DateTime Trangding_Date { get; set; }
    }
}
