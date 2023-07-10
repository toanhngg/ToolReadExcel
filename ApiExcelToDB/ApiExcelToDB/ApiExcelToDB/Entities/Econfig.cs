using System;

namespace ApiExcelToDB.Entities
{
    public class Econfig
    {
        private const string DATETIME_FORMAT_1 = "yyyy-MM-dd HH:mm:ss.fff";
        private const string DATETIME_FORMAT_2 = "yyyyMMddHHmmssfff"; // 5G
        private const string DATETIME_FORMAT_3 = "yyyy-MM-dd";
        private const string DATETIME_FORMAT_4 = "yyyyMMddHHmmss"; // bo qua millisecond
        private const string DATETIME_FORMAT_5 = "yyyyMMddHHmm"; // bo qua second
        private const string DATETIME_FORMAT_6 = "yyyyMMdd"; // bo qua hour : tao score bat dau ngay
        private const string DATETIME_FORMAT_7 = "yyyyMMdd235959999"; // hour fixed : tao score ket thuc ngay
        private const string DATETIME_FORMAT_14 = "yyMMddHHmmss"; // bo qua millisecond
        private const string DATETIME_FORMAT_15 = "yyMMddHHmm"; // bo qua second
        private const string DATETIME_FORMAT_16 = "yyMMdd"; // bo qua hour : tao score bat dau ngay
        private const string DATETIME_FORMAT_17 = "yyMMdd235959999"; // hour fixed : tao score ket thuc ngay

        public string DATETIME_FORMAT_21 = "dd.MM.yyyy";
        public string DATETIME_FORMAT_22 = "dd-MM-yyyy";
       // public string DATETIME_FORMAT_23 = "d/M/yyyy HH:mm:ss";
        public string DATETIME_FORMAT_24 = "dd-MM-yyyy"; // danh cho update date vao db

        public string ERROR = "Invalid Excel file: no worksheet found.";


        public const string __FILE_JSON_APPSETTINGS = "appsettings.json";
        static public string DateTimeNow => DateTime.Now.ToString(DATETIME_MONITOR);


        public const string DATETIME_MONITOR = DATETIME_FORMAT_1;

        public string regexRar = "Regex:RegexRar";
        public string SQLConn = "ConnectionStrings:SQLConnection";
        public string sheetName = "Setting1:SheetName";
        public string beginrow1 = "Setting1:BeginRow";
        public string beginrow2 = "Setting1:BeginRow2";
        public string beginrow3 = "Setting1:BeginRow3";
        //data1
        public string data1col1 = "Setting1:Data1:ColumOfDB";
        public string databeginRow1 = "Setting1:Data1:BeginRow";
        //2
        public string data1col2 = "Setting1:Data2:ColumOfDB";
        public string databeginRow2 = "Setting1:Data2:BeginRow";
        //3 
        public string data1col3 = "Setting1:Data3:ColumOfDB";
        public string databeginRow3 = "Setting1:Data3:BeginRow";
        //4
        public string data1col4 = "Setting1:Data4:BeginRow";
        public string databeginRow4 = "Setting1:Data4:BeginRow";
        //5
        public string data1col5 = "Setting1:Data5:ColumOfDB";
        public string databeginRow5 = "Setting1:Data5:BeginRow";
        //6
        public string data1col6 = "Setting1:Data6:ColumOfDB";
        public string databeginRow6 = "Setting1:Data6:BeginRow";
        //7
        public string data1col7 = "Setting1:Data7:ColumOfDB";
        public string databeginRow7= "Setting1:Data7:BeginRow";
        //8
        public string data1col8 = "Setting1:Data8:ColumOfDB";
        public string databeginRow8 = "Setting1:Data8:BeginRow";
        //9
        public string data1col9 = "Setting1:Data9:ColumOfDB";
        public string databeginRow9 = "Setting1:Data9:BeginRow";
        //10
        public string data1col10 = "Setting1:Data10:ColumOfDB";
        public string databeginRow10 = "Setting1:Data10:BeginRow";
        //HOSEINDEX
        public string data1col11 = "Setting1:HOSEINDEX:ColumOfDB";
        public string databeginRow11 = "Setting1:HOSEINDEX:BeginRow";
        public string databeginCell11 = "Setting1:HOSEINDEX:BeginCell";



        // data2
        public string set2beginrow1 = "Setting2:BeginRow";

        public string data2col1 = "Setting2:Data1:ColumOfDB";
        public string data2beginRow1 = "Setting2:Data1:BeginRow";
        //market
        public string marketrow1 = "MarketCap:MarketCap_1:BeginRow";

        public string marketcol = "MarketCap:MarketCap_1:ColumOfDB";
        public string marketcell1 = "MarketCap:MarketCap_1:BeginCell";

        //order matching

        public string ordermatchingSheetname = "Ordermatching:SheetName";
        public string ordermatchingcol = "Ordermatching:Ordermatching_2:ColumOfDB";

        public string ordermatchingcell = "Ordermatching:Ordermatching_2:BeginCell";
        public string ordermatchingrow = "Ordermatching:Ordermatching_2:BeginRow";
        //TCNY

        public string data2old2col = "Setting2:DataOld2:ColumOfDB";
        public string dataold2cell = "Setting2:DataOld2:BeginCell";

        public string dataold2row = "Setting2:DataOld2:BeginRow";

        //dataold3
        public string data2old3col = "Setting2:DataOld3:ColumOfDB";

        public string dataold3cell = "Setting2:DataOld3:BeginCell";
        public string dataold3row = "Setting2:DataOld3:BeginRow";
        //dataonl1

        public string data2old1col = "Setting2:DataOld1:ColumOfDB";
        public string dataold1cell = "Setting2:DataOld1:BeginCell";

        public string dataold1row = "Setting2:DataOld1:BeginRow";
        //dataold
        public string data2oldcol = "Setting2:DataOld:ColumOfDB";
        public string dataoldcell = "Setting2:DataOld:BeginCell";

        public string dataoldrow = "Setting2:DataOld:BeginRow";
        // sukien
        public string sukiencol =  "Corporate_actions:Su_kien:ColumOfDB";
        public string sukiencell = "Corporate_actions:Su_kien:BeginCell";
        public string sukienrow = "Corporate_actions:BeginRow";
        //Constituents
        public string consSheetname = "Constituents:SheetName";
        public string conscol1 = "Constituents:VNAll:ColumOfDB1";
        public string conscell1 = "Constituents:VNAll:BeginCell1";
        public string consrow = "Constituents:VNAll:BeginRow";

        public string conscol = "Constituents:VNAll:ColumOfDB";
        public string conscell = "Constituents:VNAll:BeginCell";


        public string conscolvn30 = "Constituents:VN30:ColumOfDB";
        public string conscellvn30 = "Constituents:VN30:BeginCell";
        public string consrowvn30 = "Constituents:VN30:BeginRow";

        public string conscolVNMIDCAP = "Constituents:VNMIDCAP:ColumOfDB";
        public string conscellVNMIDCAP = "Constituents:VNMIDCAP:BeginCell";
        public string consrowVNMIDCAP = "Constituents:VNMIDCAP:BeginRow";

        public string conscolVNSMALLCAP = "Constituents:VNSMALLCAP:ColumOfDB";
        public string conscellVNSMALLCAP = "Constituents:VNSMALLCAP:BeginCell";
        public string consrowVNSMALLCAP = "Constituents:VNSMALLCAP:BeginRow";

        public string conscolVN100 = "Constituents:VN100:ColumOfDB";
        public string conscellVN100 = "Constituents:VN100:BeginCell";
        public string consrowVN100 = "Constituents:VN100:BeginRow";

        public string conscolVNALLSHARE = "Constituents:VNALLSHARE:ColumOfDB";
        public string conscellVNALLSHARE = "Constituents:VNALLSHARE:BeginCell";
        public string consrowVNALLSHARE = "Constituents:VNALLSHARE:BeginRow";
        // seting 2 data1

        public string data3Sheetname = "Setting3:SheetName";
        public string data3col = "Setting3:Data1:ColumOfDB";
        public string data3cell = "Setting3:Data1:BeginRow";
        public string data3row = "Setting3:BeginRow";

        // seting 2 data2

        public string data2col = "Setting3:Data2:ColumOfDB";
        public string data2cell = "Setting3:Data2:BeginRow";
        // foreign 1 
        public string foreignsheetname = "Foreign:SheetName";
        public string foreig1ncol = "Foreign:Foreign_1:ColumOfDB";
        public string foreign1cell = "Foreign:Foreign_1:BeginCell";
        public string foreignrow = "Foreign:Foreign_1:BeginRow";
        // foreign 2 
        public string foreig2ncol = "Foreign:Foreign_2:ColumOfDB";
        public string foreign2cell2 = "Foreign:Foreign_2:BeginCell2";
        public string foreign2cell1 = "Foreign:Foreign_2:BeginCell1";
        public string foreig2nrow = "Foreign:Foreign_2:BeginRow";
        // foreign 3 
        public string foreig3ncol = "Foreign:Foreign_3:ColumOfDB";
        public string foreign3cell = "Foreign:Foreign_3:BeginCell";
        public string foreign3row = "Foreign:Foreign_3:BeginRow";

        // foreign 5
        public string foreign5col = "Foreign:Foreign_5:ColumOfDB";
        public string foreign5cell = "Foreign:Foreign_5:BeginCell";
        public string foreign5row = "Foreign:Foreign_5:BeginRow";
        // foreign cw
        public string foreigncwcol = "Foreign:Foreign_CW:ColumOfDB";
        public string foreigncwcell = "Foreign:Foreign_CW:BeginCell";
        public string foreigncwrow = "Foreign:Foreign_CW:BeginRow";
        // foreign odd
        public string foreignoddcol = "Foreign:Foreign_ODD:ColumOfDB";
        public string foreignoddcell = "Foreign:Foreign_ODD:BeginCell";
        public string foreignoddrow = "Foreign:Foreign_ODD:BeginRow";
        // foreign 6
        public string foreign6col = "Foreign:Foreign_6:ColumOfDB";
        public string foreign6cell = "Foreign:Foreign_6:BeginCell";
        public string foreign6row = "Foreign:Foreign_6:BeginRow";
        // foreign Proprietary
        public string ProprietarySheetName = "Proprietary:SheetName";
        public string ProprietaryCol = "Proprietary:Proprietary_Summary:ColumOfDB";
        public string ProprietaryCell = "Proprietary:Proprietary_Summary:BeginCell";
        public string ProprietaryRow = "Proprietary:Proprietary_Summary:BeginRow";

        // foreign Proprietary_Details
        public string Proprietary_DetailsyCol = "Proprietary:Proprietary_Details:ColumOfDB";
        public string Proprietary_DetailsCell = "Proprietary:Proprietary_Details:BeginCell";
        public string Proprietary_DetailsRow = "Proprietary:Proprietary_Details:BeginRow";
        // foreign Proprietary_DetailsCW
        public string Proprietary_DetailsCWCol = "Proprietary:Proprietary_DetailsCW:ColumOfDB";
        public string Proprietary_DetailsCWCell = "Proprietary:Proprietary_DetailsCW:BeginCell";
        public string Proprietary_DetailsCWRow = "Proprietary:Proprietary_DetailsCW:BeginRow";

        // foreign Proprietary_Order
        public string Proprietary_OrderCol = "Proprietary:Proprietary_Order:ColumOfDB";
        public string Proprietary_OrderCell = "Proprietary:Proprietary_Order:BeginCell";
        public string Proprietary_OrderRow = "Proprietary:Proprietary_Order:BeginRow";

        // foreign Proprietary_OrderCW
        public string Proprietary_OrderCWCol = "Proprietary:Proprietary_OrderCW:ColumOfDB";
        public string Proprietary_OrderCWCell = "Proprietary:Proprietary_OrderCW:BeginCell";
        public string Proprietary_OrderCWRow = "Proprietary:Proprietary_OrderCW:BeginRow";

        // foreign Trading_Result_0
        public string Trading_Result_0Row1 = "Trading_Result:Trading_Result_0:BeginRow1";
        public string Trading_Result_0Row = "Trading_Result:Trading_Result_0:BeginRow";
        public string Trading_Result_0Col = "Trading_Result:Trading_Result_0:ColumOfDB";
        public string Trading_Result_0Cell = "Trading_Result:Trading_Result_0:BeginCell";


        // foreign Trading_Result_1
        public string Trading_Result_1Row1 = "Trading_Result:Trading_Result_1:BeginRow1";
        public string Trading_Result_1Row = "Trading_Result:Trading_Result_1:BeginRow";
        public string Trading_Result_1Col = "Trading_Result:Trading_Result_1:ColumOfDB";
        public string Trading_Result_1Cell = "Trading_Result:Trading_Result_1:BeginCell";



        // foreign Trading_Result_2
        public string Trading_Result_2Row1 = "Trading_Result:Trading_Result_2:BeginRow1";
        public string Trading_Result_2Row = "Trading_Result:Trading_Result_2:BeginRow";
        public string Trading_Result_2Col = "Trading_Result:Trading_Result_2:ColumOfDB";
        public string Trading_Result_2Cell = "Trading_Result:Trading_Result_2:BeginCell";


        // foreign Trading_Result_3
        public string Trading_Result_3Row1 = "Trading_Result:Trading_Result_3:BeginRow1";
        public string Trading_Result_3Row = "Trading_Result:Trading_Result_3:BeginRow";
        public string Trading_Result_3Col = "Trading_Result:Trading_Result_3:ColumOfDB";
        public string Trading_Result_3Cell = "Trading_Result:Trading_Result_3:BeginCell";
        // foreign Trading_Result_4
        public string Trading_Result_4Row1 = "Trading_Result:Trading_Result_4:BeginRow1";
        public string Trading_Result_4Row = "Trading_Result:Trading_Result_4:BeginRow";
        public string Trading_Result_4Col = "Trading_Result:Trading_Result_4:ColumOfDB";
        public string Trading_Result_4Cell = "Trading_Result:Trading_Result_4:BeginCell";

        // foreign PutThough_Treasury PT
        public string PutThough_TreasurySheetName = "PutThough_Treasury:SheetName";
        public string PTcol1 = "PutThough_Treasury:PT:ColumOfDB1";
        public string PTcol2 = "PutThough_Treasury:PT:ColumOfDB2";
        public string PTCell = "PutThough_Treasury:PT:BeginCell";
        public string PTRow = "PutThough_Treasury:PT:BeginRow";

        // foreign PutThough_Treasury MBL
        public string MBLcol = "PutThough_Treasury:MBL:ColumOfDB";
        public string MBLCell = "PutThough_Treasury:MBL:BeginCell";
        public string MBLRow = "PutThough_Treasury:MBL:BeginRow";

        // setup date
        public string fromDate1 = "SettingDate:FromDate1";
        public string toDate1 = "SettingDate:ToDate1";
        public string fromDate2 = "SettingDate:FromDate2";
        public string toDate2 = "SettingDate:ToDate2";
        public string fromDate3 = "SettingDate:FromDate3";
        public string toDate3 = "SettingDate:ToDate3";
        public string fromDate4 = "SettingDate:FromDate4";
        public string toDate4 = "SettingDate:ToDate4";
        public string fromDate5 = "SettingDate:FromDate5";
        public string toDate5 = "SettingDate:ToDate5";
        public string fromDate6 = "SettingDate:FromDate6";
        public string toDate6 = "SettingDate:ToDate6";
        public string fromDate7 = "SettingDate:FromDate7";
        public string toDate7 = "SettingDate:ToDate7";
        public string fromDate8 = "SettingDate:FromDate8";
        public string toDate8 = "SettingDate:ToDate8";
        public string fromDate9 = "SettingDate:FromDate9";
        public string toDate9 = "SettingDate:ToDate9";
        public string fromDate10 = "SettingDate:FromDate10";
        public string toDate10 = "SettingDate:ToDate10";
        public string fromDate11 = "SettingDate:FromDate11";
        public string toDate11 = "SettingDate:ToDate11";
        public string fromDate12 = "SettingDate:FromDate12";
        public string toDate12 = "SettingDate:ToDate12";
        public string fromDate13 = "SettingDate:FromDate13";
        public string toDate13 = "SettingDate:ToDate13";
        public string fromDate14 = "SettingDate:FromDate14";
        public string toDate14 = "SettingDate:ToDate14";


        // tên file
        public string Session1 = "TKGD tung phien (Sessions).xls";
        public string Session2 = "TK Giao dich tung phien.xls";

        public string Basic1 = "TK Chi so co ban cua CP (Basic idicators).xls";
        public string Basic2 = "TK Chi so co ban cua CP (Basic Idicators).xls";

        public string MarketCap = "Thong ke giao dich - Market Cap - Turnover";

        public string CC = "CC";

        public string CP = "Chi so tai chinh cua CP.xls";
        public string TCNY = "Chi so tai chinh co ban cua TCNY";

        public string Indices = "TK hang ngay cac chi so (Indices Values).xls";

        public string Corporate = "TK su kien doanh nghiep (Corporate actions).xls";

        public string Constituents = "TK ty trong cp VNAll (Constituents VNAllshare).xls";

        public string Order1 = "TK Cung cau hang ngay (Order placement).xls";
        public string Order2 = "TK Cung cau hang ngay (Order Placement).xls";
        public string Order3 = "TK Cung cau hang ngay OM.xls";

        public string Foreign1 = "NDTNN (Foreign Trading).xls";
        public string Foreign2 = "TKGD NDTNN (Foreign trade).xls";
        public string Foreign3 = "gd cua nha DTNN.xls";
        public string Foreign4 = "TK Giao dich cua NDTNN (Sum of trade foreign investors).xls";

        public string Proprietary = "TKGD Tu doanh (Proprietary Trading).xls";

        public string CK_GDLT = "ket qua giao dich theo tung loai CK_GDLT";
        public string TradingSummary = "Tong hop KQGD (Trading Summary).xls";

        public string PutThrough = "GD thoa thuan & GD CPQ (Put-Through & Treasury).xls";
        public string PT = "PT";

        // têm bảng
        //"ClientConnect:Address"
        public string tableSession1 = "Setting1:Data1:TableName"; // "Stock_HCM_Sessions_1";
        public string tableSession1CW = "Setting1:Data2:TableName"; //"Stock_HCM_Sessions_1CW";
        public string tableSession2 = "Setting1:Data3:TableName"; //"Stock_HCM_Sessions_2";
        public string tableSession2CW = "Setting1:Data4:TableName"; // "Stock_HCM_Sessions_2CW";
        public string tableSession2ODD = "Setting1:Data5:TableName"; //"Stock_HCM_Sessions_2ODD_OM";
        public string tableSession3 = "Setting1:Data6:TableName"; // "Stock_HCM_Sessions_3";
        public string tableSession3CW = "Setting1:Data7:TableName"; // "Stock_HCM_Sessions_3CW";
        public string tableSession4 = "Setting1:Data8:TableName"; //"Stock_HCM_Sessions_4";
        public string tableSession4CW = "Setting1:Data9:TableName"; //"Stock_HCM_Sessions_4CW";
        public string tableSession4ODD = "Setting1:Data10:TableName"; //"Stock_HCM_Sessions_4ODD_PT";
        public string tableSessionHOSE = "Setting1:HOSEINDEX:TableName"; //"Stock_HCM_Sessions_HOSEINDEX";
        public string tableBasic = "Setting2:Data1:TableName"; //"Stock_HCM_Basic_idicators";
        public string tableBasic1 = "Setting2:DataOld:TableName"; //"Stock_HCM_Basic_idicators_1";
        public string tableOrder1 = "Setting3:Data1:TableName"; //"Stock_HCM_OrderPlacement";
        public string tableOrderCW = "Setting3:Data2:TableName"; //"Stock_HCM_OrderPlacement_CW";
        public string tableOrderODD = "Setting3:Data3:TableName"; //"Stock_HCM_OrderPlacement_ODD";
        public string tableForeign1 = "Foreign:Foreign_1:TableName"; //"Stock_HCM_Foreign_1";
        public string tableForeign2 = "Foreign:Foreign_2:TableName"; //"Stock_HCM_Foreign_2";
        public string tableForeign3 = "Foreign:Foreign_3:TableName"; //"Stock_HCM_Foreign_3";
        public string tableForeignCW = "Foreign:Foreign_CW:TableName"; //"Stock_HCM_Foreign_CW";
        public string tableForeignODD = "Foreign:Foreign_ODD:TableName"; //"Stock_HCM_Foreign_ODD";
        public string tableForeign5 = "Foreign:Foreign_5:TableName"; //"Stock_HCM_Foreign_5";
        public string tableProprietary1 = "Proprietary:Proprietary_Summary:TableName"; //"Stock_HCM_Proprietary_Summary";
        public string tableProprietary2 = "Proprietary:Proprietary_Details:TableName"; //"Stock_HCM_Proprietary_Details";
        public string tableProprietary3 = "Proprietary:Proprietary_DetailsCW:TableName"; //"Stock_HCM_Proprietary_Details_CW";
        public string tableProprietary4 = "Proprietary:Proprietary_Order:TableName"; //"Stock_HCM_Proprietary_Order";
        public string tableProprietary5 = "Proprietary:Proprietary_OrderCW:TableName"; //"Stock_HCM_Proprietary_Order_CW";
        public string tableCorporate = "Corporate_actions:Su_kien:TableName"; //"Stock_HCM_Corporate";
        public string tableConstituents1 = "Constituents:VNAll:TableName"; //"Stock_HCM_Constituents_VNAll";
        public string tableConstituents2 = "Constituents:VN30:TableName"; //"Stock_HCM_Constituents";
        public string tableTrading = "Trading_Result:Trading_Result_0:TableName"; //"Stock_HCM_Trading_Result";
        public string tableTrading1 = "Trading_Result:Trading_Result_1:TableName"; //"Stock_HCM_Trading_Result_1";
        public string tableTrading2 = "Trading_Result:Trading_Result_2:TableName"; //"Stock_HCM_Trading_Result_2";
        public string tableTrading4 = "Trading_Result:Trading_Result_4:TableName"; //"Stock_HCM_Trading_Result_4";
        public string tableTotalTrading = "PutThough_Treasury:PT:TableName"; //"Stock_HCM_Total_Trading";
        public string tableMarket = "MarketCap:MarketCap_1:TableName"; //"Stock_HCM_MaketCap";
        public string tableOrder = "Ordermatching:Ordermatching_2:TableName"; //"Stock_HCM_OrderMatching";
        // query insert
        public string insertTbl = "INSERT INTO";
        public string valueTbl = "VALUES";
    }
}
