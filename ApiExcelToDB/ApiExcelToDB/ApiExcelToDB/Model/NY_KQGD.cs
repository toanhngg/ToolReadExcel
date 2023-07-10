using System;
namespace ApiExcelToDB.Model
{
    public class NY_KQGD
    {
        //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighestPrice,LowestPrice,
        //Diem,PhanTram,KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT,KLGD_LL,GTGD_LL,KLGD_TC
        //,TyTrong1,GTGD_TC,TyTrong2,KLCP_LuuHanh,GTVHTT_GT,GTVHTT_TT,VDL,Trangding_Date
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double BasicPrice { get; set; }
        public double OpenPrice { get; set; }
        public double ClosePrice { get; set; }
        public double HighestPrice { get; set; }
        public double LowestPrice { get; set; }
        public double Diem { get; set; }
        public double PhanTram { get; set; }

        public double KLGD_KL { get; set; }
        public double GTGD_KL { get; set; }
        public double KLGD_TT { get; set; }
        public double GTGD_TT { get; set; }
        public double KLGD_LL { get; set; }
        public double GTGD_LL { get; set; }
        public double KLGD_TC { get; set; }
        public double TyTrong1 { get; set; }

        public double GTGD_TC { get; set; }
        public double TyTrong2 { get; set; }
        public double KLCP_LuuHanh { get; set; }
        public double GTVHTT_GT { get; set; }
        public double GTVHTT_TT { get; set; }
        public double VDL { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
