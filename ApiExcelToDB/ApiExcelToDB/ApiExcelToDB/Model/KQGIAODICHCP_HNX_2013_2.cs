using System;
namespace ApiExcelToDB.Model
{
    public class KQGIAODICHCP_HNX_2013_2
    {
      //KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT,KLGD_LL
        //,GTGD_LL,KLGD_TC,TITRONG1,GTGD_TC,TITRONG2,KLCPLH,GTVHTT_GT,GTVHTT_TT,VonDL,Trangding_Date"
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double BasicPrice { get; set; }
        public double OpenPrice { get; set; }
        public double ClosePrice { get; set; }
        public double HighestPrice { get; set; }
        public double LowestPrice { get; set; }
        public double GiaCoSo { get; set; }
        public double TDDiem { get; set; }

        public double TDPhanTram { get; set; }
        public double KLGD_KL { get; set; }
        public double GTGD_KL { get; set; }
        public double KLGD_TT { get; set; }
        public double GTGD_TT { get; set; }
      
        public double KLGD_TC { get; set; }
        public double TITRONG1 { get; set; }

        public double GTGD_TC { get; set; }
        public double TITRONG2 { get; set; }
        public double KLCPLH { get; set; }
        public double GTVHTT_GT { get; set; }
        public double GTVHTT_TT { get; set; }
        public double VonDL { get; set; }
     
        public DateTime Trangding_Date { get; set; }
    }
}
