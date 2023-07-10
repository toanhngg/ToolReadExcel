using System;
namespace ApiExcelToDB.Model
{
    public class Top10_CPGTVHMAX_HNX
    {
        //Symbol,ClosePrice,KLGD,GTVHTT,Trangding_Date
        public string Symbol { get; set; }
        public double ClosePrice { get; set; }
        public double KLGD { get; set; }
        public double GTVHTT { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
