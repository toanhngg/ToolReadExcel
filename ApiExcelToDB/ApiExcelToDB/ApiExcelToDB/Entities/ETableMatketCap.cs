using System;

namespace ApiExcelToDB.Entities
{
    public class ETableMatketCap
    {
      public DateTime TransDate{ get; set; }
      public DateTime CreateDate{ get; set; }
      public string StockCode{ get; set; }
      public float Session_Year_Count{ get; set; }
      public float TradingVolume_Year{ get; set; }
      public float AvgSession1{ get; set; }
      public float TradingValue_Year{ get; set; }
      public float AvgSession2{ get; set; }
      public float Price{ get; set; }
      public float KLNY{ get; set; }
      public float KLNY_Current{ get; set; }
      public float GTVH{ get; set; }
      public float SpeedChange{ get; set; }
    }
}
