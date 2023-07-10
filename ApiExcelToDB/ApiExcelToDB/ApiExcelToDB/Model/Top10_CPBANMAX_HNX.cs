using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPBANMAX_HNX
    {
        //Symbol,KLBAN,GTBAN,KLNG,Trangding_Date
        public string Symbol { get; set; }
        public double KLBAN { get; set; }
        public double GTBAN { get; set; }
        public double KLNG { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
