using System;
namespace ApiExcelToDB.Model
{
    public class GDTP_NDTNN
    {
        //Symbol,KLMua_KL,KLBan_KL,KL_ChenhLech,GTMua_KL,GTBan_KL,KLMua_TT,
        //KLBan_TT,KL_ChenhLech_TT,GTMua_TT,GTBan_TT,KLMua_TC,KLBan_TC,KL_ChenhLech_TC,GTMua_TC,GTBan_TC,Trangding_Date

        public string Symbol { get; set; }
      
        public double KLMua_KL { get; set; }
        public double KLBan_KL { get; set; }
        public double KL_ChenhLech { get; set; }
        public double GTMua_KL { get; set; }
        public double GTBan_KL { get; set; }

        public double KLMua_TT { get; set; }
        public double KLBan_TT { get; set; }
        public double KL_ChenhLech_TT { get; set; }
        public double GTMua_TT { get; set; }
        public double GTBan_TT { get; set; }

        public double KLMua_TC { get; set; }
        public double KLBan_TC { get; set; }
        public double KL_ChenhLech_TC { get; set; }
        public double GTMua_TC { get; set; }
        public double GTBan_TC { get; set; }

        public DateTime Trangding_Date { get; set; }
    }
}
