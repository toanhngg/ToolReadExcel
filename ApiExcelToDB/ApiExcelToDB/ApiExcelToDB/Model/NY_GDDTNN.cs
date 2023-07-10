using System;
namespace ApiExcelToDB.Model
{
    public class NY_GDDTNN
    {
      
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double KLMUA_KL { get; set; }
        public double GTMUA_KL { get; set; }
        public double KLBAN_KL { get; set; }
        public double GTBAN_KL { get; set; }
        public double KLMUA_TT { get; set; }
        public double GTMUA_TT { get; set; }
        public double KLBAN_TT { get; set; }

        public double GTBAN_TT { get; set; }
        public double KLMUA_TC { get; set; }
        public double GTMUA_TC { get; set; }
        public double GTBAN_TC { get; set; }
        public double KLBAN_TC { get; set; }
        public double KLCK_MAX { get; set; }
        public double KLCK_NDTNN { get; set; }
        public double KLCK_CDPNG { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
