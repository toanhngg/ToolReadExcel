using System;
namespace ApiExcelToDB.Model
{
    public class NDTNN2011
    {
      
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double KLCKMAX { get; set; }
        public double KLMUA_QT { get; set; }
        public double GTMUA_QT { get; set; }
        public double KLBAN_QT { get; set; }
        public double GIATRI_QT { get; set; }
        public double KLMUA_NT { get; set; }
        public double GTMUA_NT { get; set; }
        public double KLBAN_NT { get; set; }
        public double GIATRI_NT { get; set; }
        public double CurrentRoom { get; set; }
        public double KLLH { get; set; }
        public double NamGiuMax { get; set; }
        public double KLNDTN { get; set; }
      

        //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,KLBAN_QT,GIATRI_QT,KLMUA_NT,
        //GTMUA_NT,KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
        public DateTime Trangding_Date { get; set; }
    }
}
