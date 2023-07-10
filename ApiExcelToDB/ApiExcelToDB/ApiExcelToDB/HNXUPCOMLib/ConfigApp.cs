namespace ApiExcelToDB.HNXUPCOMLib
{
    public class ConfigApp
    {
        public const string EOD_2 = "UPCoM_EOD_2_Tong_hop_giao_dich_toan_thi_truong_UPCoM_A";
        public const string EOD1 = "UpCoM_EOD1_ket_qua_giao_dich_co_phieu_dkgd_A";
        public const string EOD4 = "UPCom_EOD4_Thong_ke_cung_cau_thi_truong_co_phieu_dkgd_A";

        public const string EOD5 = "UpCom_EOD5_Thong_tin_co_ban_ve_cong_ty_dkgd_A";
        public const string EOD6 = "UPCom_EOD6_Giao_dich_nha_dau_tu_nuoc_ngoai_co_phieu_dkgd_A";
        public const string EOD7 = "UPCoM_EOD7_gia_giao_dich_ngay_ke_tiep_co_phieu_dkgd_A";

        public const string NY_EOD_2 = "NY_EOD_2_Tong_hop_giao_dich_toan_thi_truong_Niem_yet_A";
        public const string NY_EOD1 = "NY_EOD1_ket_qua_giao_dich_co_phieu_niem_yet_A";
        public const string NY_EOD4 = "NY_EOD4_Thong_ke_cung_cau_thi_truong_co_phieu_niem_yet_A";

        public const string NY_EOD5 = "NY_EOD5_Thong_tin_co_ban_ve_cong_ty_niem_yet_A";
        public const string NY_EOD6 = "NY_EOD6_Giao_dich_nha_dau_tu_nuoc_ngoai_co_phieu_niem_yet_A";
        public const string NY_EOD7 = "NY_EOD7_gia_giao_dich_ngay_ke_tiep_co_phieu_niem_yet_A";

        public const string NY_21 = "NY_2.1";
        public const string NY_22 = "NY_2.2";
        public const string NY_23 = "NY_2.3";
        public const string NY_24 = "NY_2.4";
        public const string NY_25 = "NY_2.5";

        public const string NY_1 = "2.1";
        public const string NY_2 = "2.2";
        public const string NY_3 = "2.3";
        public const string NY_4 = "2.4";
        public const string NY_5 = "2.5";

        public const string CONNECT_STRING = "Configs";
        public string ConnectionString { get; set; }
        public string folderPath { get; set; }
        public string FormDate { get; set; }
        public string DateNew { get; set; }
        public string ToDate { get; set; }
        public string Regex { get; set; }

        public string Regex_2 { get; set; }
        public string Regex_3 { get; set; }

        public string Regex_4 { get; set; }
        public string Regex_5 { get; set; }
        public string Regex_6 { get; set; }
        public string Regex_7 { get; set; }
        public DataViewT6 UPCoM_GDNDTNN_2011 { get; set; }
        public DataFile2011 HNX_File_2011 { get; set; }
        public DataFile2010 HNX_File_2010 { get; set; }
        public DataView01 NY_GDNDTNN_Phien { get; set; }
        public DataView01 NY_KQGD_Phien { get; set; }
        public DataView01 NY_KQGD_Phien2 { get; set; }
        public DataView01 NY_EOD4_View { get; set; }
        public DataView01 NY_CPNY_Phien { get; set; }
        public DataView02_NY NY_TT_Phien { get; set; }
        /// <summary>
        /// /
        /// </summary>
        /// // upcom 2011
        public DataView01 UPCoM_GDNDTNN_Phien { get; set; }
        public DataView01 UPCoM_KQGD_Phien { get; set; }
        public DataView01 UPCoM_TKCC { get; set; }
        public DataView01 UPCoM_CPDKGD_Phien { get; set; }
        public DataView02 UPCoM_TT_Phien { get; set; }


        /// 
        public DataView01 NY21 { get; set; }
        public DataView01 NY23 { get; set; }
        public DataView01 NY24 { get; set; }
        public DataView01 NY25 { get; set; }
        public DataView02_NY NY22 { get; set; }
        public DataView02 DataEOD2 { get; set; }
        public DataView01EDO1 DataEOD1 { get; set; }
        public DataView01 DataEOD4 { get; set; }
        public DataView01 DataEOD5 { get; set; }
        public DataView01 DataEOD6 { get; set; }
        public DataView01 DataEOD7 { get; set; }

        public DataView02_NY NY_DataEOD2 { get; set; }
        public DataView01EDO1 NY_DataEOD1 { get; set; }
        public DataView01 NY_DataEOD4 { get; set; }
        public DataView01 NY_DataEOD5 { get; set; }
        public DataView01 NY_DataEOD6 { get; set; }
        public DataView01 NY_DataEOD7 { get; set; }

        public struct DataView01
        {
            public string FileName { get; set; }

            public string SheetName { get; set; }
            public string TableName { get; set; }
            public string SPName { get; set; }

            public string BeginCell { get; set; }
            public string Column { get; set; }

        }
        public struct DataViewHNX2011
        {

            public string SheetName { get; set; }
            public string TableName { get; set; }
            public string SPName { get; set; }

            public string BeginCell { get; set; }
            public string Column { get; set; }

        }
        public struct HNX2011
        {
            public string TableName { get; set; }
            public string SPName { get; set; }

            public string BeginCell { get; set; }
            public string Column { get; set; }
        }
        public struct DataViewHNX2011_2
        {

            public string SheetName { get; set; }

            public HNX2011 Top10CK_GTGDL { get; set; }
            public HNX2011 Top10CK_KLGDL { get; set; }
            public HNX2011 Top10CP_GTNYL { get; set; }
            public HNX2011 Top10CK_TANGGIA { get; set; }
            public HNX2011 Top10CK_GIAMGIA { get; set; }

            public HNX2011 Chi_Tieu_2011 { get; set; }
            public HNX2011 Top10CK_NDTNN { get; set; }
            public HNX2011 KLGD_TOP2011_MR { get; set; }
            public HNX2011 GTGD_TOP2011_MR { get; set; }
            public HNX2011 TangGiam_TOP2011_MR { get; set; }

            public HNX2011 CKNTDNN_TOP2011_MR { get; set; }

        }
        public struct DataViewHNX2010_2
        {

            public string SheetName { get; set; }

            public HNX2011 Top10CK_GTGDL { get; set; }
            public HNX2011 Top10CK_KLGDL { get; set; }
            public HNX2011 Top10CP_CLGMAX { get; set; }
            public HNX2011 Top10CK_TANGGIA { get; set; }
            public HNX2011 Top10CK_GIAMGIA { get; set; }



        }
        public struct DataFile2011
        {
            public DataViewHNX2011 TT_DKGD_2011 { get; set; }

            public DataViewHNX2011 TH_DATLENH_2011 { get; set; }
            public DataViewHNX2011 NDTNN_2011 { get; set; }
            public DataViewHNX2011 KQGD_2011 { get; set; }

            public DataViewHNX2011 KQGDTH_2011 { get; set; }
            public DataViewHNX2011_2 Top_2011 { get; set; }

        }
        public struct DataFile2010
        {
            public DataViewHNX2011 GD_TRAIPHIEU { get; set; }
            public DataViewHNX2011 GDTP_NDTNN { get; set; }
            public DataViewHNX2010_2 Top_2010 { get; set; }

        }
        public struct DataView01EDO1
        {
            public string FileName { get; set; }

            public string SheetName { get; set; }

            public DataViewMin1 Data1 { get; set; }
            public DataViewMin1 Data2 { get; set; }
        }
        public struct DataViewT6
        {
            public string SheetName { get; set; }

            public string TableName { get; set; }

            public string BeginCell { get; set; }
            public string Column { get; set; }
        }



        public struct DataView02
        {
            public string FileName { get; set; }

            public string SheetName { get; set; }
            public DataViewMin1 Data_Table_Chi_Tieu { get; set; }
            public DataViewMin1 Data_Table_Top10_CPGDT { get; set; }
            public DataViewMin1 Data_Table_Top10_CPTPRICE { get; set; }
            public DataViewMin1 Data_Table_Top10_KLGDM { get; set; }
            public DataViewMin1 Data_Table_Top10_CPGIAMGIA { get; set; }

        }
        public struct DataView02_NY
        {
            public string FileName { get; set; }

            public string SheetName { get; set; }
            public DataViewMin1 Data_Table_Chi_Tieu_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPGDMAX_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPNYGTMAX_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPMUAMAX_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPTANGPRICE_HNX { get; set; }

            public DataViewMin1 Data_Table_Top10_KLGDMAX_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPGTVHMAX_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPBANMAX_HNX { get; set; }
            public DataViewMin1 Data_Table_Top10_CPGIAMPRICE_HNX { get; set; }

        }
        public struct DataViewMin1
        {
            public string TableName { get; set; }
            public string SPName { get; set; }

            public string BeginCell { get; set; }
            public string Column { get; set; }

        }



    }
}
