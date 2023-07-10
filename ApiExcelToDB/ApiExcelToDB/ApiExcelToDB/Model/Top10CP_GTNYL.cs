using System;
namespace ApiExcelToDB.Model
{
    public class Top10CP_GTNYL
    {
      
        public string Symbol { get; set; }
      
        public double AvePrice { get; set; }
        public double Volume { get; set; }
        public double GiaTriNY { get; set; }
      
        public DateTime Trangding_Date { get; set; }
    }
}
