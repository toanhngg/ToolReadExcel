using System;
namespace ApiExcelToDB.Model
{
    public class Top10CK_KLGDL
    {
      
        public string Symbol { get; set; }
      
        public double AvePrice { get; set; }
        public double Volume { get; set; }
        public double PhanTram { get; set; }
        public double WeightN { get; set; }
      
        public DateTime Trangding_Date { get; set; }
    }
}
