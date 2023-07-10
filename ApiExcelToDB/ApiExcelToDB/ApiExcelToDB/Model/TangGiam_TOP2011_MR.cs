using System;
namespace ApiExcelToDB.Model
{
    public class TangGiam_TOP2011_MR
    {
      
        public string Symbol { get; set; }
      
        public double AvePrice { get; set; }
        public double MucTang { get; set; }
        public double PTTangGiam { get; set; }

        public double KLGD { get; set; }
        public double CEILINGPRICE { get; set; }
        public double ChenhLechTran { get; set; }
        public double FLOORPRICES { get; set; }
        public double ChenhLechSan { get; set; }
     
        public DateTime Trangding_Date { get; set; }
    }
}
