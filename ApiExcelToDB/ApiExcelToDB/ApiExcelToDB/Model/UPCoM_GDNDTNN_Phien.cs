using System;
namespace ApiExcelToDB.Model
{
    public class UPCoM_GDNDTNN_Phien
    {
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double KLMUA_KL { get; set; }
        public double GTMUA_KL { get; set; }
        public double KLBAN_KL { get; set; }

        public double GTBAN_KL { get; set; }
        public double KLMUA_TT { get; set; }
        public double GTMUA_TT { get; set; }

      //  public double KLCK_NDTNN { get; set; }
        public double GTBAN_TT { get; set; }
        public double KLBAN_TT { get; set; }
        public double KLMUA_TC { get; set; }
        public double GTMUA_TC { get; set; }
        public double KLBAN_TC { get; set; }
        public double GTBAN_TC { get; set; }
        public double KLCK_MAX { get; set; }
        public double KLCK_NDTNN { get; set; }
        public double KLCK_CDPNG { get; set; }
        //STT,Symbol,KLMUA_KL,GTMUA_KL,KLBAN_KL,
        //GTBAN_KL,KLMUA_TT,GTMUA_TT,KLBAN_TT,GTBAN_TT,
        //KLMUA_TC,GTMUA_TC,KLBAN_TC,GTBAN_TC,KLCK_MAX,KLCK_NDTNN,KLCK_CDPNG,Trangding_Date
        public DateTime Trangding_Date { get; set; }
    }
}
