using System;
namespace ApiExcelToDB.Model
{
    public class CKNTDNN_TOP2011_MR
    {
        
        public string Symbol { get; set; }
        public double KLMua { get; set; }
        public double GTMua { get; set; }
        public double KLBan { get; set; }
        public double GTBan { get; set; }
        public double KLDPNamGiu { get; set; }
        //Symbol,KLMua,GTMua,KLDPNamGiu,Trangding_Date
        public DateTime Trangding_Date { get; set; }
    }
}
