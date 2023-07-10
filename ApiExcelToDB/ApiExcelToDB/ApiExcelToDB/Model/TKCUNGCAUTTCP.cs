using System;
namespace ApiExcelToDB.Model
{
    public class TKCUNGCAUTTCP
    {
      
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double SLDATMUA_KL { get; set; }
        public double KLDATMUA_KL { get; set; }
        public double SLDATBAN_KL { get; set; }
        public double KLDATBAN_KL { get; set; }
        public double SLDATMUA_TT { get; set; }
        public double KLDATMUA_TT { get; set; }
        public double SLDATBAN_TT { get; set; }

        public double KLDATBAN_TT { get; set; }
        public double SLDATMUA_TC { get; set; }
        public double KLDATMUA_TC { get; set; }
        public double SLDATBAN_TC { get; set; }
        public double KLDATBAN_TC { get; set; }
        public double KLDUMUA { get; set; }
        public double KLDUBAN { get; set; }
        public double KLTHUCHIEN { get; set; }
        public double GTTHUCHIEN { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
