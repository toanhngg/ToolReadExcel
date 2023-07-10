using System;
namespace ApiExcelToDB.Model
{
    public class GTGD_TOP2011_MR
    {
      
        public string Symbol { get; set; }
      
        public double AvePrice { get; set; }
        public double KLGD { get; set; }
        public double KLNY { get; set; }
        public double GTNY_Trieu { get; set; }
        public double GTNY_Dong { get; set; }
     
        public DateTime Trangding_Date { get; set; }
    }
}
