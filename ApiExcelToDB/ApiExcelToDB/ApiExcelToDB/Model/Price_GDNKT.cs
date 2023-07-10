using System;
namespace ApiExcelToDB.Model
{
    public class Price_GDNKT
    {
        //STT,Symbol,Market,BasicPrice_HT,CeilingPrice_HT,FloorPrice_HT,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
        public int STT { get; set; }
        public string Symbol { get; set; }
        public string Market { get; set; }
        public double BasicPrice_HT { get; set; }
        public double CeilingPrice_HT { get; set; }
        public double FloorPrice_HT { get; set; }
        public double BasicPrice_KT { get; set; }
        public double CeilingPrice_KT { get; set; }
        public double FloorPrice_KT { get; set; }
        public DateTime Trangding_Date { get; set; }
    }
}
