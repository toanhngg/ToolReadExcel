using System;

namespace ApiExcelToDB.Model
{
    public class UPCoM_TKCC
    {
        public int STT { get; set; }
        public string Symbol { get; set; }
        public double SLDATMUA { get; set; }
        public double KLDATMUA { get; set; }
        public double SLDATBAN { get; set; }

        public double KLDATBAN { get; set; }
        public double CLMUABAN { get; set; }
      
        public DateTime Trangding_Date { get; set; }
    }
}
