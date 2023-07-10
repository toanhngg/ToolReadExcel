using System;
namespace ApiExcelToDB.Model
{
    public class KQGIAODICHCP2
    {
     
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double BasicPrice { get; set; }
        public double OpenPrice { get; set; }
        public double HighestPrice { get; set; }
        public double LowestPrice { get; set; }
        public double AveragePrice { get; set; }
        public double TDDiem { get; set; }

        public double TDPhanTram { get; set; }
        public double KLGDC_KL { get; set; }
        public double KLGDL_KL { get; set; }
        public double GTGDC_KL { get; set; }
        public double GTGDL_KL { get; set; }
        public double KLGDC_TT { get; set; }
        public double KLGDL_TT { get; set; }
        public double GTGDC_TT { get; set; }
        public double GTGDL_TT { get; set; }

        public double KLGD_TC { get; set; }
        public double TITRONG1 { get; set; }
        public double GTGD_TC { get; set; }
        public double TITRONG2 { get; set; }
        public double KLCPLH { get; set; }
        public double GTVHTT_GT { get; set; }
        public double GTVHTT_TT { get; set; }

        public double TrangThaiCK { get; set; }
        public double TinhTrangCK { get; set; }
        public double TrangThaiThucHienQuyen { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
