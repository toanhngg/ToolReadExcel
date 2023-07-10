using System;
namespace ApiExcelToDB.Model
{
    public class NY_TTCP_2013
    {
        //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2014,Co_Tuc_2015,PE,
        //EPS2015,ROE2015,ROA2015,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double KLCP_NY { get; set; }
        public double KLCP_LH { get; set; }
        public double Co_Tuc_2013 { get; set; }
        public double Co_Tuc_2014 { get; set; }
        public double PE { get; set; }
        public double EPS2014 { get; set; }
        public double ROE2014 { get; set; }
        public double ROA2014 { get; set; }
        public double BasicPrice_KT { get; set; }
        public double CeilingPrice_KT { get; set; }
        public double FloorPrice_KT { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
