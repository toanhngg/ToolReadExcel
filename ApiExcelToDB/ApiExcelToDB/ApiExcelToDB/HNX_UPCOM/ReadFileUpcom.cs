using ApiExcelToDB.HNXUPCOMLib;
using ApiExcelToDB.Model;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
namespace ApiExcelToDB.HNX_UPCOM
{
    public class ReadFileUpcom
    {
        private readonly IConfiguration _config;
        private readonly ConfigApp configs;
        private readonly ConfigTable configTable;


        public ReadFileUpcom(IConfiguration configuration,string name)
        {
            _config = configuration;
            configs = new ConfigApp();
            configTable = new ConfigTable(configs);
            _config.GetSection(ConfigApp.CONNECT_STRING).Bind(configs);
            YouMi(name);
        }

        public void YouMi(string name)
        {
            // Chuỗi cần kiểm tra

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            // Biểu thức chính quy (regex pattern)
            Regex regexPattern = new Regex(configs.Regex);
            Regex regexPattern_2 = new Regex(configs.Regex_2);
            Regex regexPattern_3 = new Regex(configs.Regex_3);
            Regex regexPattern_4 = new Regex(configs.Regex_4);

            Regex regexPattern_5 = new Regex(configs.Regex_5);
            Regex regexPattern_6 = new Regex(configs.Regex_6);
            Regex regexPattern_7 = new Regex(configs.Regex_7);
            string pattern = @"UPCoM_KQGD_Phien_(\d{4}\.\d{2}\.\d{2})\. \((.*?)\)";

            //dùng chung
            StringBuilder mssqlBuilder_HNX = new StringBuilder();

            //THÔNG TIN CƠ BẢN VỀ CÔNG TY ĐKGD
            string folderPath = configs.folderPath + name;

            var filePaths = Directory.GetFiles(folderPath, "*.xls", SearchOption.AllDirectories);

            foreach (var filePath in filePaths)
            {

                string[] parts = filePath.Split('\\');

                // Lấy phần cuối cùng
                string lastPart = parts[parts.Length - 1];
                Match match = regexPattern.Match(lastPart);
                Match match_2 = regexPattern_2.Match(lastPart);
                Match match_3 = regexPattern_3.Match(lastPart);
                Match match_4 = regexPattern_4.Match(lastPart);
                // Match match_4_1 = regexPattern_4_1.Match(filePath);

                Match match_4_1 = Regex.Match(lastPart, pattern);
                Match match_5 = regexPattern_5.Match(lastPart);
                Match match_6 = regexPattern_6.Match(lastPart);
                Match match_7 = regexPattern_7.Match(lastPart);
                if (match.Success)
                {
                    string dateString = match.Groups[1].Value;
                    string thitruong = match.Groups[2].Value;
                    string maFile = match.Groups[3].Value;

                    DateTime date = DateTime.ParseExact(dateString, "yyyyMMdd", CultureInfo.InvariantCulture);
                    string dateS = date.ToString("yyyy-MM-dd");
                    DateTime dateFile = DateTime.ParseExact(dateS, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    DateTime dateNew = DateTime.ParseExact(configs.DateNew, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                    switch (thitruong)
                    {

                        case ConfigApp.EOD5:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTen(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.DataEOD5.SheetName];

                                        string[] column = configs.DataEOD5.BeginCell.Split(',');
                                        for (int i = 11; i < dataTable.Rows.Count - 5; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                THONGTINCB ttcb = new THONGTINCB();

                                                ttcb.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                ttcb.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    ttcb.PriceCloseAverage = 0;

                                                }
                                                else { ttcb.PriceCloseAverage = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    ttcb.KLCPNY = 0;

                                                }
                                                else { ttcb.KLCPNY = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    ttcb.KLCPLH = 0;

                                                }
                                                else { ttcb.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    ttcb.EPS = 0;

                                                }
                                                else
                                                {
                                                    ttcb.EPS = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    ttcb.ROE = 0;

                                                }
                                                else { ttcb.ROE = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    ttcb.ROA = 0;

                                                }
                                                else { ttcb.ROA = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    ttcb.GTTT = 0;

                                                }
                                                else { ttcb.GTTT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                ttcb.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM(ttcb, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD5.TableName + EDalResult.__STRING_VALUES + eBulkScript.MssqlScript.TrimEnd(',') + ";");
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            configTable.ExecBulkScript(mssqlBuilder_HNX.ToString());
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.EOD6:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenEDO6(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.DataEOD6.SheetName];

                                        string[] column = configs.DataEOD6.BeginCell.Split(',');
                                        for (int i = 10; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                GIAODICHNHADAUTUNN gdndtnn = new GIAODICHNHADAUTUNN();

                                                gdndtnn.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                gdndtnn.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    gdndtnn.KLMUA_KL = 0;

                                                }
                                                else { gdndtnn.KLMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    gdndtnn.GTMUA_KL = 0;

                                                }
                                                else { gdndtnn.GTMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    gdndtnn.KLBAN_KL = 0;

                                                }
                                                else { gdndtnn.KLBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    gdndtnn.GTBAN_KL = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn.GTBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    gdndtnn.KLMUA_TT = 0;

                                                }
                                                else { gdndtnn.KLMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    gdndtnn.GTMUA_TT = 0;

                                                }
                                                else { gdndtnn.GTMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    gdndtnn.KLBAN_TT = 0;

                                                }
                                                else { gdndtnn.KLBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    gdndtnn.GTBAN_TT = 0;

                                                }
                                                else { gdndtnn.GTBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    gdndtnn.KLMUA_TC = 0;

                                                }
                                                else { gdndtnn.KLMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    gdndtnn.GTMUA_TC = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn.GTMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    gdndtnn.KLBAN_TC = 0;

                                                }
                                                else { gdndtnn.KLBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    gdndtnn.GTBAN_TC = 0;

                                                }
                                                else { gdndtnn.GTBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    gdndtnn.KLCK_MAX = 0;

                                                }
                                                else { gdndtnn.KLCK_MAX = Convert.ToDouble(dataTable.Rows[i][column[14]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    gdndtnn.KLCK_NDTNN = 0;

                                                }
                                                else { gdndtnn.KLCK_NDTNN = Convert.ToDouble(dataTable.Rows[i][column[15]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    gdndtnn.KLCK_CDPNG = 0;

                                                }
                                                else { gdndtnn.KLCK_CDPNG = Convert.ToDouble(dataTable.Rows[i][column[16]]); }
                                                gdndtnn.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, gdndtnn, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD6.TableName + EDalResult.__STRING_VALUES + eBulkScript.MssqlScript.TrimEnd(',') + ";");
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            configTable.ExecBulkScript(mssqlBuilder_HNX.ToString());
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.EOD4:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenEDO4(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.DataEOD4.SheetName];

                                        string[] column = configs.DataEOD4.BeginCell.Split(',');
                                        for (int i = 10; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                TKCUNGCAUTTCP cc = new TKCUNGCAUTTCP();


                                                cc.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc.SLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc.SLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc.KLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc.KLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc.SLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc.SLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc.KLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc.KLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc.SLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc.SLDATMUA_TT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc.KLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc.KLDATMUA_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc.SLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc.SLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc.KLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc.KLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc.SLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc.SLDATMUA_TC = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc.KLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc.KLDATMUA_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {

                                                    cc.SLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                }
                                                else { cc.SLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {

                                                    cc.KLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                }
                                                else { cc.KLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {

                                                    cc.KLDUMUA = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                }
                                                else { cc.KLDUMUA = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {

                                                    cc.KLDUBAN = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                }
                                                else { cc.KLDUBAN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {

                                                    cc.KLTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                }
                                                else { cc.KLTHUCHIEN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                {

                                                    cc.GTTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                }
                                                else { cc.GTTHUCHIEN = 0; }
                                                cc.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, cc, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD4.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.EOD1:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        if (dateFile < dateNew)
                                        {
                                            var dataSetX = configTable.DatTenEDO1S(dataSet);
                                            float view;
                                            DataTable dataTable = dataSetX.Tables[configs.DataEOD1.SheetName];

                                            string[] column = configs.DataEOD1.Data2.BeginCell.Split(',');
                                            for (int i = 8; i < dataTable.Rows.Count - 16; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {

                                                    KQGIAODICHCP2 kq = new KQGIAODICHCP2();


                                                    kq.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    kq.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {

                                                        kq.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                    }
                                                    else { kq.BasicPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {

                                                        kq.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                    }
                                                    else { kq.OpenPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {

                                                        kq.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                    }
                                                    else { kq.HighestPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {

                                                        kq.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                    }
                                                    else { kq.LowestPrice = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {

                                                        kq.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                    }
                                                    else { kq.AveragePrice = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {

                                                        kq.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                    }
                                                    else { kq.TDDiem = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {

                                                        kq.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                    }
                                                    else { kq.TDPhanTram = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {

                                                        kq.KLGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                    }
                                                    else { kq.KLGDC_KL = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {

                                                        kq.KLGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                    }
                                                    else { kq.KLGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                    {

                                                        kq.GTGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                    }
                                                    else { kq.GTGDC_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                    {

                                                        kq.GTGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                    }
                                                    else { kq.GTGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                    {

                                                        kq.KLGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                    }
                                                    else { kq.KLGDC_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                    {

                                                        kq.KLGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                    }
                                                    else { kq.KLGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                    {

                                                        kq.GTGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                    }
                                                    else { kq.GTGDC_TT = 0; }



                                                    if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                    {

                                                        kq.GTGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                    }
                                                    else { kq.GTGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                    {

                                                        kq.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                    }
                                                    else { kq.KLGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                    {

                                                        kq.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                    }
                                                    else { kq.TITRONG1 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                    {

                                                        kq.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                    }
                                                    else { kq.GTGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                    {

                                                        kq.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                    }
                                                    else { kq.TITRONG2 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                    {

                                                        kq.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                                    }
                                                    else { kq.KLCPLH = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[22]].ToString(), out view))
                                                    {

                                                        kq.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[22]]);

                                                    }
                                                    else { kq.GTVHTT_GT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[23]].ToString(), out view))
                                                    {

                                                        kq.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[23]]);

                                                    }
                                                    else { kq.GTVHTT_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[24]].ToString(), out view))
                                                    {

                                                        kq.TrangThaiCK = Convert.ToDouble(dataTable.Rows[i][column[24]]);

                                                    }
                                                    else { kq.TrangThaiCK = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[25]].ToString(), out view))
                                                    {

                                                        kq.TinhTrangCK = Convert.ToDouble(dataTable.Rows[i][column[25]]);

                                                    }
                                                    else { kq.TinhTrangCK = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[26]].ToString(), out view))
                                                    {

                                                        kq.TrangThaiThucHienQuyen = Convert.ToDouble(dataTable.Rows[i][column[26]]);

                                                    }
                                                    else { kq.TrangThaiThucHienQuyen = 0; }
                                                    kq.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, null, null, null, null, kq);

                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD1.Data2.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();
                                            }
                                        }
                                        else
                                        {
                                            var dataSetX = configTable.DatTenEDO1(dataSet);
                                            float view;
                                            DataTable dataTable = dataSetX.Tables[configs.DataEOD1.SheetName];

                                            string[] column = configs.DataEOD1.Data1.BeginCell.Split(',');
                                            for (int i = 8; i < dataTable.Rows.Count - 1; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {

                                                    KQGIAODICHCP kq = new KQGIAODICHCP();


                                                    kq.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    kq.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {

                                                        kq.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                    }
                                                    else { kq.BasicPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {

                                                        kq.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                    }
                                                    else { kq.OpenPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {

                                                        kq.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                    }
                                                    else { kq.HighestPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {

                                                        kq.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                    }
                                                    else { kq.LowestPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {

                                                        kq.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                    }
                                                    else { kq.ClosePrice = 0; }


                                                    if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {

                                                        kq.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                    }
                                                    else { kq.AveragePrice = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {

                                                        kq.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                    }
                                                    else { kq.TDDiem = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {

                                                        kq.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                    }
                                                    else { kq.TDPhanTram = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {

                                                        kq.KLGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                    }
                                                    else { kq.KLGDC_KL = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                    {

                                                        kq.KLGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                    }
                                                    else { kq.KLGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                    {

                                                        kq.GTGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                    }
                                                    else { kq.GTGDC_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                    {

                                                        kq.GTGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                    }
                                                    else { kq.GTGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                    {

                                                        kq.KLGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                    }
                                                    else { kq.KLGDC_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                    {

                                                        kq.KLGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                    }
                                                    else { kq.KLGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                    {

                                                        kq.GTGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                    }
                                                    else { kq.GTGDC_TT = 0; }



                                                    if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                    {

                                                        kq.GTGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                    }
                                                    else { kq.GTGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                    {

                                                        kq.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                    }
                                                    else { kq.KLGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                    {

                                                        kq.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                    }
                                                    else { kq.TITRONG1 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                    {

                                                        kq.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                    }
                                                    else { kq.GTGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                    {

                                                        kq.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                                    }
                                                    else { kq.TITRONG2 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[22]].ToString(), out view))
                                                    {

                                                        kq.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[22]]);

                                                    }
                                                    else { kq.KLCPLH = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[23]].ToString(), out view))
                                                    {

                                                        kq.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[23]]);

                                                    }
                                                    else { kq.GTVHTT_GT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[24]].ToString(), out view))
                                                    {

                                                        kq.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[24]]);

                                                    }
                                                    else { kq.GTVHTT_TT = 0; }
                                                    kq.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, kq, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD1.Data1.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();
                                            }
                                        }



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.EOD7:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenEDO7(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.DataEOD7.SheetName];

                                        string[] column = configs.DataEOD7.BeginCell.Split(',');
                                        for (int i = 8; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                Price_GDNKT price = new Price_GDNKT();


                                                price.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                price.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                //STT,Symbol,Market,BasicPrice_HT,CeilingPrice_HT,FloorPrice_HT,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
                                                price.Market = dataTable.Rows[i][column[2]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    price.BasicPrice_HT = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { price.BasicPrice_HT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    price.CeilingPrice_HT = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { price.CeilingPrice_HT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    price.FloorPrice_HT = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { price.FloorPrice_HT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    price.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { price.BasicPrice_KT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    price.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { price.CeilingPrice_KT = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    price.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { price.FloorPrice_KT = 0; }

                                                price.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, null, null, null, price, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD7.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.EOD_2:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX_1 = configTable.DatTenEDO2_1(dataSet);
                                        var dataSetX_2 = configTable.DatTenEDO2_2(dataSet);

                                        float view;
                                        DataTable dataTable = dataSetX_1.Tables[configs.DataEOD2.SheetName];
                                        DataTable dataTable2 = dataSetX_2.Tables[configs.DataEOD2.SheetName];

                                        string[] column = configs.DataEOD2.Data_Table_Chi_Tieu.BeginCell.Split(',');
                                        string[] column_2 = configs.DataEOD2.Data_Table_Top10_CPGDT.BeginCell.Split(',');
                                        string[] column_3 = configs.DataEOD2.Data_Table_Top10_CPTPRICE.BeginCell.Split(',');
                                        string[] column_4 = configs.DataEOD2.Data_Table_Top10_KLGDM.BeginCell.Split(',');
                                        string[] column_5 = configs.DataEOD2.Data_Table_Top10_CPGIAMGIA.BeginCell.Split(',');
                                        for (int i = 4; i < dataTable.Rows.Count - 10; i++)
                                        {
                                            Chi_Tieu_UPCOM ct = new Chi_Tieu_UPCOM();

                                            //Chi_Tieu,Don_Vi,So_Lieu,Trangding_Date

                                            ct.Chi_Tieu = dataTable.Rows[i][column[0]].ToString();
                                            ct.Don_Vi = dataTable.Rows[i][column[1]].ToString();
                                            if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {

                                                ct.So_Lieu = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                            }
                                            else { ct.So_Lieu = 0; }

                                            ct.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, ct, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD2.Data_Table_Chi_Tieu.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        for (int i = 6; i < dataTable2.Rows.Count - 15; i++)
                                        {
                                            Top10_CPGDT cpgdt = new Top10_CPGDT();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpgdt.Symbol = dataTable2.Rows[i][column_2[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_2[1]].ToString(), out view))
                                            {
                                                cpgdt.GTGD = Convert.ToDouble(dataTable2.Rows[i][column_2[1]]);
                                            }
                                            else { cpgdt.GTGD = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[2]].ToString(), out view))
                                            {

                                                cpgdt.TyTrong = Convert.ToDouble(dataTable2.Rows[i][column_2[2]]);

                                            }
                                            else { cpgdt.TyTrong = 0; }

                                            cpgdt.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, cpgdt, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD2.Data_Table_Top10_CPGDT.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_3 = configTable.DatTenEDO2_3(dataSet);
                                        DataTable dataTable3 = dataSetX_3.Tables[configs.DataEOD2.SheetName];

                                        for (int i = 21; i < dataTable3.Rows.Count - 0; i++)
                                        {
                                            Top10_CPTPRICE cptprice = new Top10_CPTPRICE();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cptprice.Symbol = dataTable3.Rows[i][column_3[0]].ToString();
                                            if (float.TryParse(dataTable3.Rows[i][column_3[1]].ToString(), out view))
                                            {
                                                cptprice.MucTang = Convert.ToDouble(dataTable3.Rows[i][column_3[1]]);
                                            }
                                            else { cptprice.MucTang = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[2]].ToString(), out view))
                                            {

                                                cptprice.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column_3[2]]);

                                            }
                                            else { cptprice.TyLeTang = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[3]].ToString(), out view))
                                            {

                                                cptprice.KLGD = Convert.ToDouble(dataTable3.Rows[i][column_3[3]]);

                                            }
                                            else { cptprice.KLGD = 0; }

                                            cptprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, cptprice, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD2.Data_Table_Top10_CPTPRICE.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_4 = configTable.DatTenEDO2_4(dataSet);
                                        DataTable dataTable4 = dataSetX_4.Tables[configs.DataEOD2.SheetName];
                                        for (int i = 6; i < dataTable4.Rows.Count - 15; i++)
                                        {
                                            Top10_KLGDM klgdm = new Top10_KLGDM();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            klgdm.Symbol = dataTable4.Rows[i][column_4[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_4[1]].ToString(), out view))
                                            {
                                                klgdm.KLGD = Convert.ToDouble(dataTable4.Rows[i][column_4[1]]);
                                            }
                                            else { klgdm.KLGD = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[2]].ToString(), out view))
                                            {

                                                klgdm.TyTrong = Convert.ToDouble(dataTable4.Rows[i][column_4[2]]);

                                            }
                                            else { klgdm.TyTrong = 0; }

                                            klgdm.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, null, klgdm, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD2.Data_Table_Top10_KLGDM.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                        var dataSetX_5 = configTable.DatTenEDO2_5(dataSet);
                                        DataTable dataTable5 = dataSetX_5.Tables[configs.DataEOD2.SheetName];

                                        for (int i = 21; i < dataTable5.Rows.Count - 0; i++)
                                        {
                                            Top10_CPGIAMGIA cpgiamgia = new Top10_CPGIAMGIA();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgiamgia.Symbol = dataTable5.Rows[i][column_5[0]].ToString();
                                            if (float.TryParse(dataTable5.Rows[i][column_5[1]].ToString(), out view))
                                            {
                                                cpgiamgia.MucGIAM = Convert.ToDouble(dataTable5.Rows[i][column_5[1]]);
                                            }
                                            else { cpgiamgia.MucGIAM = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[2]].ToString(), out view))
                                            {

                                                cpgiamgia.TyLeGiam = Convert.ToDouble(dataTable5.Rows[i][column_5[2]]);

                                            }
                                            else { cpgiamgia.TyLeGiam = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[3]].ToString(), out view))
                                            {

                                                cpgiamgia.KLGD = Convert.ToDouble(dataTable5.Rows[i][column_5[3]]);

                                            }
                                            else { cpgiamgia.KLGD = 0; }

                                            cpgiamgia.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, null, null, cpgiamgia, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.DataEOD2.Data_Table_Top10_CPGIAMGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_EOD5:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();

                                        var dataSetX = configTable.DatTenHNX(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY_DataEOD5.SheetName];
                                        if (dataTable != null)
                                        {
                                            string[] column = configs.NY_DataEOD5.BeginCell.Split(',');
                                            for (int i = 12; i < dataTable.Rows.Count - 6; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {
                                                    THONGTINCB_HNX ttcb_hnx = new THONGTINCB_HNX();

                                                    ttcb_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    ttcb_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.PriceCloseAverage = 0;

                                                    }
                                                    else { ttcb_hnx.PriceCloseAverage = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.KLCPNY = 0;

                                                    }
                                                    else { ttcb_hnx.KLCPNY = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.KLCPLH = 0;

                                                    }
                                                    else { ttcb_hnx.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.EPS = 0;

                                                    }
                                                    else
                                                    {
                                                        ttcb_hnx.EPS = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                    }
                                                    if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.EPS4 = 0;

                                                    }
                                                    else
                                                    {
                                                        ttcb_hnx.EPS4 = Convert.ToDouble(dataTable.Rows[i][column[6]]);
                                                    }
                                                    if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.PE = 0;

                                                    }
                                                    else
                                                    {
                                                        ttcb_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[7]]);
                                                    }
                                                    if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.ROE = 0;

                                                    }
                                                    else { ttcb_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.ROA = 0;

                                                    }
                                                    else { ttcb_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {
                                                        ttcb_hnx.GTTT = 0;

                                                    }
                                                    else { ttcb_hnx.GTTT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                    ttcb_hnx.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX(ttcb_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD5.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);

                                                mssqlBuilder_HNX.Clear();
                                            }

                                        }
                                    }


                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_EOD6:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenEDO6(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY_DataEOD6.SheetName];

                                        string[] column = configs.NY_DataEOD6.BeginCell.Split(',');
                                        for (int i = 10; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {

                                                GIAODICHNHADAUTUNN_HNX gdndtnn_hnx = new GIAODICHNHADAUTUNN_HNX();

                                                gdndtnn_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                gdndtnn_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLMUA_KL = 0;

                                                }
                                                else { gdndtnn_hnx.KLMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTMUA_KL = 0;

                                                }
                                                else { gdndtnn_hnx.GTMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLBAN_KL = 0;

                                                }
                                                else { gdndtnn_hnx.KLBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTBAN_KL = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn_hnx.GTBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLMUA_TT = 0;

                                                }
                                                else { gdndtnn_hnx.KLMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTMUA_TT = 0;

                                                }
                                                else { gdndtnn_hnx.GTMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLBAN_TT = 0;

                                                }
                                                else { gdndtnn_hnx.KLBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTBAN_TT = 0;

                                                }
                                                else { gdndtnn_hnx.GTBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLMUA_TC = 0;

                                                }
                                                else { gdndtnn_hnx.KLMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTMUA_TC = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn_hnx.GTMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLBAN_TC = 0;

                                                }
                                                else { gdndtnn_hnx.KLBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTBAN_TC = 0;

                                                }
                                                else { gdndtnn_hnx.GTBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLCK_MAX = 0;

                                                }
                                                else { gdndtnn_hnx.KLCK_MAX = Convert.ToDouble(dataTable.Rows[i][column[14]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLCK_NDTNN = 0;

                                                }
                                                else { gdndtnn_hnx.KLCK_NDTNN = Convert.ToDouble(dataTable.Rows[i][column[15]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLCK_CDPNG = 0;

                                                }
                                                else { gdndtnn_hnx.KLCK_CDPNG = Convert.ToDouble(dataTable.Rows[i][column[16]]); }
                                                gdndtnn_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX(null, gdndtnn_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD6.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_EOD4:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenEDO4(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY_DataEOD4.SheetName];

                                        string[] column = configs.NY_DataEOD4.BeginCell.Split(',');
                                        for (int i = 10; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {

                                                TKCUNGCAUTTCP_HNX cc_hnx = new TKCUNGCAUTTCP_HNX();


                                                cc_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc_hnx.SLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc_hnx.KLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc_hnx.SLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc_hnx.KLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc_hnx.SLDATMUA_TT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc_hnx.KLDATMUA_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc_hnx.SLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc_hnx.KLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc_hnx.SLDATMUA_TC = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc_hnx.KLDATMUA_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                }
                                                else { cc_hnx.SLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                }
                                                else { cc_hnx.KLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDUMUA = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                }
                                                else { cc_hnx.KLDUMUA = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDUBAN = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                }
                                                else { cc_hnx.KLDUBAN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {

                                                    cc_hnx.KLTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                }
                                                else { cc_hnx.KLTHUCHIEN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                {

                                                    cc_hnx.GTTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                }
                                                else { cc_hnx.GTTHUCHIEN = 0; }
                                                cc_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, cc_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD4.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_EOD7:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenEDO7(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY_DataEOD7.SheetName];

                                        string[] column = configs.NY_DataEOD7.BeginCell.Split(',');
                                        for (int i = 8; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {

                                                Price_GDNKT_HNX price_hnx = new Price_GDNKT_HNX();


                                                price_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                price_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                //STT,Symbol,Market,Basicprice_hnx_HT,Ceilingprice_hnx_HT,Floorprice_hnx_HT,Basicprice_hnx_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
                                                price_hnx.Market = dataTable.Rows[i][column[2]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    price_hnx.BasicPrice_HT = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { price_hnx.BasicPrice_HT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    price_hnx.CeilingPrice_HT = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { price_hnx.CeilingPrice_HT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    price_hnx.FloorPrice_HT = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { price_hnx.FloorPrice_HT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    price_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { price_hnx.BasicPrice_KT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    price_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { price_hnx.CeilingPrice_KT = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    price_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { price_hnx.FloorPrice_KT = 0; }

                                                price_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, price_hnx, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD7.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_EOD1:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();

                                        if (dateFile < dateNew)
                                        {
                                            var dataSetX = configTable.DatTenEDO1_HNXS(dataSet);
                                            float view;
                                            DataTable dataTable = dataSetX.Tables[configs.NY_DataEOD1.SheetName];

                                            string[] column = configs.NY_DataEOD1.Data2.BeginCell.Split(',');
                                            for (int i = 8; i < dataTable.Rows.Count - 16; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {

                                                    KQGIAODICHCP_HNX2 kq_hnx = new KQGIAODICHCP_HNX2();


                                                    kq_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    kq_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {

                                                        kq_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                    }
                                                    else { kq_hnx.BasicPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {

                                                        kq_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                    }
                                                    else { kq_hnx.OpenPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {

                                                        kq_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                    }
                                                    else { kq_hnx.ClosePrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {

                                                        kq_hnx.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                    }
                                                    else { kq_hnx.HighestPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {

                                                        kq_hnx.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                    }
                                                    else { kq_hnx.LowestPrice = 0; }


                                                    if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {

                                                        kq_hnx.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                    }
                                                    else { kq_hnx.TDDiem = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {

                                                        kq_hnx.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                    }
                                                    else { kq_hnx.TDPhanTram = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                    }
                                                    else { kq_hnx.KLGDC_KL = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                    }
                                                    else { kq_hnx.KLGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                    }
                                                    else { kq_hnx.GTGDC_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                    }
                                                    else { kq_hnx.GTGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                    }
                                                    else { kq_hnx.KLGDC_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                    }
                                                    else { kq_hnx.KLGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                    }
                                                    else { kq_hnx.GTGDC_TT = 0; }



                                                    if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                    }
                                                    else { kq_hnx.GTGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                    }
                                                    else { kq_hnx.KLGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                    {

                                                        kq_hnx.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                    }
                                                    else { kq_hnx.TITRONG1 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                    }
                                                    else { kq_hnx.GTGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                    {

                                                        kq_hnx.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                    }
                                                    else { kq_hnx.TITRONG2 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                                    }
                                                    else { kq_hnx.KLCPLH = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[22]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[22]]);

                                                    }
                                                    else { kq_hnx.GTVHTT_GT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[23]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[23]]);

                                                    }
                                                    else { kq_hnx.GTVHTT_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[24]].ToString(), out view))
                                                    {

                                                        kq_hnx.TrangThaiCK = Convert.ToDouble(dataTable.Rows[i][column[24]]);

                                                    }
                                                    else { kq_hnx.TrangThaiCK = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[25]].ToString(), out view))
                                                    {

                                                        kq_hnx.TinhTrangCK = Convert.ToDouble(dataTable.Rows[i][column[25]]);

                                                    }
                                                    else { kq_hnx.TinhTrangCK = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[26]].ToString(), out view))
                                                    {

                                                        kq_hnx.TrangThaiThucHienQuyen = Convert.ToDouble(dataTable.Rows[i][column[26]]);

                                                    }
                                                    else { kq_hnx.TrangThaiThucHienQuyen = 0; }
                                                    kq_hnx.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, null, null, kq_hnx);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD1.Data2.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                        }
                                        else
                                        {
                                            var dataSetX = configTable.DatTenEDO1_HNX(dataSet);
                                            float view;
                                            DataTable dataTable = dataSetX.Tables[configs.NY_DataEOD1.SheetName];

                                            string[] column = configs.NY_DataEOD1.Data1.BeginCell.Split(',');
                                            for (int i = 8; i < dataTable.Rows.Count - 1; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {

                                                    KQGIAODICHCP_HNX kq_hnx = new KQGIAODICHCP_HNX();


                                                    kq_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    kq_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {

                                                        kq_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                    }
                                                    else { kq_hnx.BasicPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {

                                                        kq_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                    }
                                                    else { kq_hnx.OpenPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {

                                                        kq_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                    }
                                                    else { kq_hnx.ClosePrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {

                                                        kq_hnx.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                    }
                                                    else { kq_hnx.HighestPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {

                                                        kq_hnx.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                    }
                                                    else { kq_hnx.LowestPrice = 0; }


                                                    if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {

                                                        kq_hnx.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                    }
                                                    else { kq_hnx.TDDiem = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {

                                                        kq_hnx.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                    }
                                                    else { kq_hnx.TDPhanTram = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                    }
                                                    else { kq_hnx.KLGDC_KL = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                    }
                                                    else { kq_hnx.KLGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDC_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                    }
                                                    else { kq_hnx.GTGDC_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDL_KL = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                    }
                                                    else { kq_hnx.GTGDL_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                    }
                                                    else { kq_hnx.KLGDC_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                    }
                                                    else { kq_hnx.KLGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDC_TT = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                    }
                                                    else { kq_hnx.GTGDC_TT = 0; }



                                                    if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGDL_TT = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                    }
                                                    else { kq_hnx.GTGDL_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                    }
                                                    else { kq_hnx.KLGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                    {

                                                        kq_hnx.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                    }
                                                    else { kq_hnx.TITRONG1 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                    }
                                                    else { kq_hnx.GTGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                    {

                                                        kq_hnx.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                    }
                                                    else { kq_hnx.TITRONG2 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                                    }
                                                    else { kq_hnx.KLCPLH = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[22]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[22]]);

                                                    }
                                                    else { kq_hnx.GTVHTT_GT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[23]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[23]]);

                                                    }
                                                    else { kq_hnx.GTVHTT_TT = 0; }
                                                    kq_hnx.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, kq_hnx, null, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                }
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD1.Data1.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }
                                            }






                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_EOD_2:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX_1 = configTable.DatTenEDO2_1(dataSet);
                                        var dataSetX_2 = configTable.DatTenEDO2_2_HNX(dataSet);

                                        float view;
                                        DataTable dataTable = dataSetX_1.Tables[configs.NY_DataEOD2.SheetName];
                                        DataTable dataTable2 = dataSetX_2.Tables[configs.NY_DataEOD2.SheetName];

                                        string[] column = configs.NY_DataEOD2.Data_Table_Chi_Tieu_HNX.BeginCell.Split(',');
                                        string[] column_2 = configs.NY_DataEOD2.Data_Table_Top10_CPGDMAX_HNX.BeginCell.Split(',');
                                        string[] column_3 = configs.NY_DataEOD2.Data_Table_Top10_CPNYGTMAX_HNX.BeginCell.Split(',');
                                        string[] column_4 = configs.NY_DataEOD2.Data_Table_Top10_CPMUAMAX_HNX.BeginCell.Split(',');
                                        string[] column_5 = configs.NY_DataEOD2.Data_Table_Top10_CPTANGPRICE_HNX.BeginCell.Split(',');

                                        string[] column_6 = configs.NY_DataEOD2.Data_Table_Top10_KLGDMAX_HNX.BeginCell.Split(',');
                                        string[] column_7 = configs.NY_DataEOD2.Data_Table_Top10_CPGTVHMAX_HNX.BeginCell.Split(',');
                                        string[] column_8 = configs.NY_DataEOD2.Data_Table_Top10_CPBANMAX_HNX.BeginCell.Split(',');
                                        string[] column_9 = configs.NY_DataEOD2.Data_Table_Top10_CPGIAMPRICE_HNX.BeginCell.Split(',');

                                        for (int i = 4; i < dataTable.Rows.Count - 34; i++)
                                        {
                                            Chi_Tieu_HNX ct = new Chi_Tieu_HNX();

                                            //Chi_Tieu,Don_Vi,So_Lieu,Trangding_Date

                                            ct.Chi_Tieu = dataTable.Rows[i][column[0]].ToString();
                                            ct.Don_Vi = dataTable.Rows[i][column[1]].ToString();
                                            if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {

                                                ct.So_Lieu = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                            }
                                            else { ct.So_Lieu = 0; }

                                            ct.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, ct, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Chi_Tieu_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        for (int i = 5; i < dataTable2.Rows.Count - 46; i++)
                                        {
                                            Top10_CPGDMAX_HNX cpgdmax = new Top10_CPGDMAX_HNX();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpgdmax.Symbol = dataTable2.Rows[i][column_2[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_2[1]].ToString(), out view))
                                            {
                                                cpgdmax.ClosePrice = Convert.ToDouble(dataTable2.Rows[i][column_2[1]]);
                                            }
                                            else { cpgdmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[2]].ToString(), out view))
                                            {
                                                cpgdmax.GTGD = Convert.ToDouble(dataTable2.Rows[i][column_2[2]]);
                                            }
                                            else { cpgdmax.GTGD = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[3]].ToString(), out view))
                                            {

                                                cpgdmax.TyTrong = Convert.ToDouble(dataTable2.Rows[i][column_2[3]]);

                                            }
                                            else { cpgdmax.TyTrong = 0; }

                                            cpgdmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, cpgdmax, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPGDMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_3 = configTable.DatTenEDO2_3(dataSet);
                                        DataTable dataTable3 = dataSetX_3.Tables[configs.NY_DataEOD2.SheetName];

                                        for (int i = 21; i < dataTable3.Rows.Count - 30; i++)
                                        {
                                            Top10_CPNYGTMAX_HNX cpnygtmax = new Top10_CPNYGTMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpnygtmax.Symbol = dataTable3.Rows[i][column_3[0]].ToString();
                                            if (float.TryParse(dataTable3.Rows[i][column_3[1]].ToString(), out view))
                                            {
                                                cpnygtmax.ClosePrice = Convert.ToDouble(dataTable3.Rows[i][column_3[1]]);
                                            }
                                            else { cpnygtmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[2]].ToString(), out view))
                                            {

                                                cpnygtmax.KLGD = Convert.ToDouble(dataTable3.Rows[i][column_3[2]]);

                                            }
                                            else { cpnygtmax.KLGD = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[3]].ToString(), out view))
                                            {

                                                cpnygtmax.GTNY = Convert.ToDouble(dataTable3.Rows[i][column_3[3]]);

                                            }
                                            else { cpnygtmax.GTNY = 0; }

                                            cpnygtmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, cpnygtmax, null, null, null, null, null, null, null);
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPNYGTMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_4 = configTable.DatTenEDO2_4_HNX(dataSet);
                                        DataTable dataTable4 = dataSetX_4.Tables[configs.NY_DataEOD2.SheetName];
                                        for (int i = 36; i < dataTable4.Rows.Count - 15; i++)
                                        {
                                            Top10_CPMUAMAX_HNX cpmuamax = new Top10_CPMUAMAX_HNX();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpmuamax.Symbol = dataTable4.Rows[i][column_4[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_4[1]].ToString(), out view))
                                            {
                                                cpmuamax.KLGD = Convert.ToDouble(dataTable4.Rows[i][column_4[1]]);
                                            }
                                            else { cpmuamax.KLGD = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[2]].ToString(), out view))
                                            {

                                                cpmuamax.GTMUA = Convert.ToDouble(dataTable4.Rows[i][column_4[2]]);

                                            }
                                            else { cpmuamax.GTMUA = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[3]].ToString(), out view))
                                            {

                                                cpmuamax.KLNG = Convert.ToDouble(dataTable4.Rows[i][column_4[3]]);

                                            }
                                            else { cpmuamax.KLNG = 0; }

                                            cpmuamax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, cpmuamax, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPMUAMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_5 = configTable.DatTenEDO2_5_HNX(dataSet);
                                        DataTable dataTable5 = dataSetX_5.Tables[configs.NY_DataEOD2.SheetName];

                                        for (int i = 51; i < dataTable5.Rows.Count - 0; i++)
                                        {
                                            Top10_CPTANGPRICE_HNX cptangprice = new Top10_CPTANGPRICE_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cptangprice.Symbol = dataTable5.Rows[i][column_5[0]].ToString();
                                            if (float.TryParse(dataTable5.Rows[i][column_5[1]].ToString(), out view))
                                            {
                                                cptangprice.ClosePrice = Convert.ToDouble(dataTable5.Rows[i][column_5[1]]);
                                            }
                                            else { cptangprice.ClosePrice = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[2]].ToString(), out view))
                                            {

                                                cptangprice.MucTang = Convert.ToDouble(dataTable5.Rows[i][column_5[2]]);

                                            }
                                            else { cptangprice.MucTang = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[3]].ToString(), out view))
                                            {

                                                cptangprice.TyLeTang = Convert.ToDouble(dataTable5.Rows[i][column_5[3]]);

                                            }
                                            else { cptangprice.TyLeTang = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[4]].ToString(), out view))
                                            {

                                                cptangprice.KLGD = Convert.ToDouble(dataTable5.Rows[i][column_5[4]]);

                                            }
                                            else { cptangprice.KLGD = 0; }

                                            cptangprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, cptangprice, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPTANGPRICE_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_6 = configTable.DatTenEDO2_6_HNX(dataSet);
                                        DataTable dataTable6 = dataSetX_6.Tables[configs.NY_DataEOD2.SheetName];

                                        for (int i = 5; i < dataTable6.Rows.Count - 46; i++)
                                        {
                                            Top10_KLGDMAX_HNX klgdmax = new Top10_KLGDMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            klgdmax.Symbol = dataTable6.Rows[i][column_6[0]].ToString();
                                            if (float.TryParse(dataTable6.Rows[i][column_6[1]].ToString(), out view))
                                            {
                                                klgdmax.ClosePrice = Convert.ToDouble(dataTable6.Rows[i][column_6[1]]);
                                            }
                                            else { klgdmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable6.Rows[i][column_6[2]].ToString(), out view))
                                            {

                                                klgdmax.KLGD = Convert.ToDouble(dataTable6.Rows[i][column_6[2]]);

                                            }
                                            else { klgdmax.KLGD = 0; }
                                            if (float.TryParse(dataTable6.Rows[i][column_6[3]].ToString(), out view))
                                            {

                                                klgdmax.TyTrong = Convert.ToDouble(dataTable6.Rows[i][column_6[3]]);

                                            }
                                            else { klgdmax.TyTrong = 0; }


                                            klgdmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, klgdmax, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_KLGDMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_7 = configTable.DatTenEDO2_7_HNX(dataSet);
                                        DataTable dataTable7 = dataSetX_7.Tables[configs.NY_DataEOD2.SheetName];

                                        for (int i = 21; i < dataTable7.Rows.Count - 30; i++)
                                        {
                                            Top10_CPGTVHMAX_HNX cpgtvhmax = new Top10_CPGTVHMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgtvhmax.Symbol = dataTable7.Rows[i][column_7[0]].ToString();
                                            if (float.TryParse(dataTable7.Rows[i][column_7[1]].ToString(), out view))
                                            {
                                                cpgtvhmax.ClosePrice = Convert.ToDouble(dataTable7.Rows[i][column_7[1]]);
                                            }
                                            else { cpgtvhmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable7.Rows[i][column_7[2]].ToString(), out view))
                                            {

                                                cpgtvhmax.KLGD = Convert.ToDouble(dataTable7.Rows[i][column_7[2]]);

                                            }
                                            else { cpgtvhmax.KLGD = 0; }
                                            if (float.TryParse(dataTable7.Rows[i][column_7[3]].ToString(), out view))
                                            {

                                                cpgtvhmax.GTVHTT = Convert.ToDouble(dataTable7.Rows[i][column_7[3]]);

                                            }
                                            else { cpgtvhmax.GTVHTT = 0; }


                                            cpgtvhmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, cpgtvhmax, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPGTVHMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        var dataSetX_8 = configTable.DatTenEDO2_8_HNX(dataSet);
                                        DataTable dataTable8 = dataSetX_8.Tables[configs.NY_DataEOD2.SheetName];

                                        for (int i = 36; i < dataTable8.Rows.Count - 15; i++)
                                        {
                                            Top10_CPBANMAX_HNX cpbanmax = new Top10_CPBANMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpbanmax.Symbol = dataTable8.Rows[i][column_8[0]].ToString();
                                            if (float.TryParse(dataTable8.Rows[i][column_8[1]].ToString(), out view))
                                            {
                                                cpbanmax.KLBAN = Convert.ToDouble(dataTable8.Rows[i][column_8[1]]);
                                            }
                                            else { cpbanmax.KLBAN = 0; }
                                            if (float.TryParse(dataTable8.Rows[i][column_8[2]].ToString(), out view))
                                            {

                                                cpbanmax.GTBAN = Convert.ToDouble(dataTable8.Rows[i][column_8[2]]);

                                            }
                                            else { cpbanmax.GTBAN = 0; }
                                            if (float.TryParse(dataTable8.Rows[i][column_8[3]].ToString(), out view))
                                            {

                                                cpbanmax.KLNG = Convert.ToDouble(dataTable8.Rows[i][column_8[3]]);

                                            }
                                            else { cpbanmax.KLNG = 0; }


                                            cpbanmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, cpbanmax, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPBANMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                        var dataSetX_9 = configTable.DatTenEDO2_9_HNX(dataSet);
                                        DataTable dataTable9 = dataSetX_9.Tables[configs.NY_DataEOD2.SheetName];

                                        for (int i = 51; i < dataTable9.Rows.Count - 0; i++)
                                        {
                                            Top10_CPGIAMPRICE_HNX cpgiamprice = new Top10_CPGIAMPRICE_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgiamprice.Symbol = dataTable9.Rows[i][column_9[0]].ToString();
                                            if (float.TryParse(dataTable9.Rows[i][column_9[1]].ToString(), out view))
                                            {
                                                cpgiamprice.ClosePrice = Convert.ToDouble(dataTable9.Rows[i][column_9[1]]);
                                            }
                                            else { cpgiamprice.ClosePrice = 0; }
                                            if (float.TryParse(dataTable9.Rows[i][column_9[2]].ToString(), out view))
                                            {

                                                cpgiamprice.MucGiam = Convert.ToDouble(dataTable9.Rows[i][column_9[2]]);

                                            }
                                            else { cpgiamprice.MucGiam = 0; }
                                            if (float.TryParse(dataTable9.Rows[i][column_9[3]].ToString(), out view))
                                            {

                                                cpgiamprice.TyLeGiam = Convert.ToDouble(dataTable9.Rows[i][column_9[3]]);

                                            }
                                            else { cpgiamprice.TyLeGiam = 0; }

                                            if (float.TryParse(dataTable9.Rows[i][column_9[4]].ToString(), out view))
                                            {

                                                cpgiamprice.KLGD = Convert.ToDouble(dataTable9.Rows[i][column_9[4]]);

                                            }
                                            else { cpgiamprice.TyLeGiam = 0; }

                                            cpgiamprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, null, cpgiamprice, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }


                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_DataEOD2.Data_Table_Top10_CPGIAMPRICE_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";

                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        default:
                            // code block
                            Console.WriteLine("File chưa định dạng để read!");
                            break;
                    }
                    ///

                }
                if (match_2.Success)
                {
                    string dateString = match_2.Groups[4].Value;
                    string thitruong = match_2.Groups[1].Value + "_" + match_2.Groups[2].Value;
                    DateTime date = DateTime.ParseExact(dateString, "yyyyMMdd", CultureInfo.InvariantCulture);
                    string dateS = date.ToString("yyyy-MM-dd");
                    DateTime dateFile = DateTime.ParseExact(dateS, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    //DateTime dateNew = DateTime.ParseExact(configs.DateNew, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    switch (thitruong)
                    {
                        case ConfigApp.NY_21:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY21(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY21.SheetName];

                                        string[] column = configs.NY21.BeginCell.Split(',');
                                        for (int i = 6; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NY_KQGD kqgd = new NY_KQGD();

                                                kqgd.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                kqgd.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighestPrice,LowestPrice,

                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    kqgd.BasicPrice = 0;

                                                }
                                                else { kqgd.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    kqgd.OpenPrice = 0;

                                                }
                                                else { kqgd.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    kqgd.ClosePrice = 0;

                                                }
                                                else { kqgd.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    kqgd.HighestPrice = 0;

                                                }
                                                else
                                                {
                                                    kqgd.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    kqgd.LowestPrice = 0;

                                                }
                                                else { kqgd.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    kqgd.Diem = 0;

                                                }
                                                else { kqgd.Diem = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    kqgd.PhanTram = 0;

                                                }
                                                else { kqgd.PhanTram = Convert.ToDouble(dataTable.Rows[i][column[8]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    kqgd.KLGD_KL = 0;

                                                }
                                                else { kqgd.KLGD_KL = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    kqgd.KLGD_TT = 0;

                                                }
                                                else { kqgd.KLGD_TT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    kqgd.GTGD_TT = 0;

                                                }
                                                else
                                                {
                                                    kqgd.GTGD_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    kqgd.KLGD_LL = 0;

                                                }
                                                else { kqgd.KLGD_LL = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                //Diem,PhanTram,KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT,KLGD_LL,GTGD_LL,KLGD_TC
                                                //,TyTrong1,GTGD_TC,TyTrong2,KLCP_LuuHanh,GTVHTT_GT,GTVHTT_TT,VDL,Trangding_Date
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    kqgd.GTGD_LL = 0;

                                                }
                                                else { kqgd.GTGD_LL = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    kqgd.KLGD_TC = 0;

                                                }
                                                else { kqgd.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    kqgd.TyTrong1 = 0;

                                                }
                                                else { kqgd.TyTrong1 = Convert.ToDouble(dataTable.Rows[i][column[15]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    kqgd.GTGD_TC = 0;

                                                }
                                                else { kqgd.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[16]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                {
                                                    kqgd.TyTrong2 = 0;

                                                }
                                                else { kqgd.TyTrong2 = Convert.ToDouble(dataTable.Rows[i][column[17]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                {
                                                    kqgd.KLCP_LuuHanh = 0;

                                                }
                                                else { kqgd.KLCP_LuuHanh = Convert.ToDouble(dataTable.Rows[i][column[18]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                {
                                                    kqgd.GTVHTT_GT = 0;

                                                }
                                                else { kqgd.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[19]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                {
                                                    kqgd.GTVHTT_TT = 0;

                                                }
                                                else { kqgd.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[20]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                {
                                                    kqgd.VDL = 0;

                                                }
                                                else { kqgd.VDL = Convert.ToDouble(dataTable.Rows[i][column[21]]); }

                                                kqgd.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptNY2017(kqgd, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY21.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }



                                    }
                                }



                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;

                        case ConfigApp.NY_23:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY23(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY23.SheetName];

                                        string[] column = configs.NY23.BeginCell.Split(',');
                                        for (int i = 5; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NY_ThongKeCC cc_hnx = new NY_ThongKeCC();


                                                cc_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc_hnx.SLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc_hnx.KLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc_hnx.SLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc_hnx.KLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc_hnx.SLDATMUA_TT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc_hnx.KLDATMUA_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc_hnx.SLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc_hnx.KLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc_hnx.SLDATMUA_TC = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc_hnx.KLDATMUA_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                }
                                                else { cc_hnx.SLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                }
                                                else { cc_hnx.KLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDUMUA = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                }
                                                else { cc_hnx.KLDUMUA = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDUBAN = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                }
                                                else { cc_hnx.KLDUBAN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {

                                                    cc_hnx.KLTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                }
                                                else { cc_hnx.KLTHUCHIEN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                {

                                                    cc_hnx.GTTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                }
                                                else { cc_hnx.GTTHUCHIEN = 0; }
                                                cc_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptNY2017(null, cc_hnx, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY23.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_24:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY24(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY24.SheetName];

                                        string[] column = configs.NY24.BeginCell.Split(',');
                                        for (int i = 5; i < dataTable.Rows.Count - 6; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NY_GDDTNN gdndtnn = new NY_GDDTNN();

                                                gdndtnn.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                gdndtnn.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    gdndtnn.KLMUA_KL = 0;

                                                }
                                                else { gdndtnn.KLMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    gdndtnn.GTMUA_KL = 0;

                                                }
                                                else { gdndtnn.GTMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    gdndtnn.KLBAN_KL = 0;

                                                }
                                                else { gdndtnn.KLBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    gdndtnn.GTBAN_KL = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn.GTBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    gdndtnn.KLMUA_TT = 0;

                                                }
                                                else { gdndtnn.KLMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    gdndtnn.GTMUA_TT = 0;

                                                }
                                                else { gdndtnn.GTMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    gdndtnn.KLBAN_TT = 0;

                                                }
                                                else { gdndtnn.KLBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    gdndtnn.GTBAN_TT = 0;

                                                }
                                                else { gdndtnn.GTBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    gdndtnn.KLMUA_TC = 0;

                                                }
                                                else { gdndtnn.KLMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    gdndtnn.GTMUA_TC = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn.GTMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    gdndtnn.KLBAN_TC = 0;

                                                }
                                                else { gdndtnn.KLBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    gdndtnn.GTBAN_TC = 0;

                                                }
                                                else { gdndtnn.GTBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    gdndtnn.KLCK_MAX = 0;

                                                }
                                                else { gdndtnn.KLCK_MAX = Convert.ToDouble(dataTable.Rows[i][column[14]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    gdndtnn.KLCK_NDTNN = 0;

                                                }
                                                else { gdndtnn.KLCK_NDTNN = Convert.ToDouble(dataTable.Rows[i][column[15]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    gdndtnn.KLCK_CDPNG = 0;

                                                }
                                                else { gdndtnn.KLCK_CDPNG = Convert.ToDouble(dataTable.Rows[i][column[16]]); }
                                                gdndtnn.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptNY2017(null, null, gdndtnn, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY24.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        case ConfigApp.NY_25:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY25(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.NY25.SheetName];
                                        //  Console.WriteLine(filePath );
                                        string[] column = configs.NY25.BeginCell.Split(',');
                                        for (int i = 4; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NY_TTCP ttcb = new NY_TTCP();
                                                ttcb.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);


                                                ttcb.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2014,Co_Tuc_2015,PE,EPS2015,ROE2015,

                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    ttcb.KLCP_NY = 0;

                                                }
                                                else { ttcb.KLCP_NY = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    ttcb.KLCP_LH = 0;

                                                }
                                                else { ttcb.KLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    ttcb.Co_Tuc_2014 = 0;

                                                }
                                                else { ttcb.Co_Tuc_2014 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    ttcb.Co_Tuc_2015 = 0;

                                                }
                                                else
                                                {
                                                    ttcb.Co_Tuc_2015 = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    ttcb.PE = 0;

                                                }
                                                else { ttcb.PE = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    ttcb.EPS2015 = 0;

                                                }
                                                else { ttcb.EPS2015 = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    ttcb.ROE2015 = 0;

                                                }
                                                else { ttcb.ROE2015 = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    ttcb.ROA2015 = 0;

                                                }
                                                else { ttcb.ROA2015 = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    ttcb.BasicPrice_KT = 0;

                                                }
                                                else { ttcb.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    ttcb.CeilingPrice_KT = 0;

                                                }
                                                else { ttcb.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    ttcb.FloorPrice_KT = 0;

                                                }
                                                else { ttcb.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                //ROA2015,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
                                                ttcb.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptNY2017(null, null, null, ttcb);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY25.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();


                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        case ConfigApp.NY_22:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX_1 = configTable.DatTenNY22_1(dataSet);
                                        var dataSetX_2 = configTable.DatTenNY22_2(dataSet);

                                        float view;
                                        DataTable dataTable = dataSetX_1.Tables[configs.NY22.SheetName];
                                        DataTable dataTable2 = dataSetX_2.Tables[configs.NY22.SheetName];

                                        string[] column = configs.NY22.Data_Table_Chi_Tieu_HNX.BeginCell.Split(',');
                                        string[] column_2 = configs.NY22.Data_Table_Top10_CPGDMAX_HNX.BeginCell.Split(',');
                                        string[] column_3 = configs.NY22.Data_Table_Top10_CPNYGTMAX_HNX.BeginCell.Split(',');
                                        string[] column_4 = configs.NY22.Data_Table_Top10_CPMUAMAX_HNX.BeginCell.Split(',');
                                        string[] column_5 = configs.NY22.Data_Table_Top10_CPTANGPRICE_HNX.BeginCell.Split(',');

                                        string[] column_6 = configs.NY22.Data_Table_Top10_KLGDMAX_HNX.BeginCell.Split(',');
                                        string[] column_7 = configs.NY22.Data_Table_Top10_CPGTVHMAX_HNX.BeginCell.Split(',');
                                        string[] column_8 = configs.NY22.Data_Table_Top10_CPBANMAX_HNX.BeginCell.Split(',');
                                        string[] column_9 = configs.NY22.Data_Table_Top10_CPGIAMPRICE_HNX.BeginCell.Split(',');

                                        for (int i = 4; i < dataTable.Rows.Count - 32; i++)
                                        {
                                            Chi_Tieu_HNX ct = new Chi_Tieu_HNX();

                                            //Chi_Tieu,Don_Vi,So_Lieu,Trangding_Date

                                            ct.Chi_Tieu = dataTable.Rows[i][column[0]].ToString();
                                            // ct.Don_Vi = dataTable.Rows[i][column[1]].ToString();
                                            if (float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {

                                                ct.So_Lieu = Convert.ToDouble(dataTable.Rows[i][column[1]]);

                                            }
                                            else { ct.So_Lieu = 0; }

                                            ct.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, ct, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Chi_Tieu_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        for (int i = 5; i < dataTable2.Rows.Count - 44; i++)
                                        {
                                            Top10_CPGDMAX_HNX cpgdmax = new Top10_CPGDMAX_HNX();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpgdmax.Symbol = dataTable2.Rows[i][column_2[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_2[1]].ToString(), out view))
                                            {
                                                cpgdmax.ClosePrice = Convert.ToDouble(dataTable2.Rows[i][column_2[1]]);
                                            }
                                            else { cpgdmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[2]].ToString(), out view))
                                            {
                                                cpgdmax.GTGD = Convert.ToDouble(dataTable2.Rows[i][column_2[2]]);
                                            }
                                            else { cpgdmax.GTGD = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[3]].ToString(), out view))
                                            {

                                                cpgdmax.TyTrong = Convert.ToDouble(dataTable2.Rows[i][column_2[3]]);

                                            }
                                            else { cpgdmax.TyTrong = 0; }

                                            cpgdmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, cpgdmax, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPGDMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_3 = configTable.DatTenNY22_3(dataSet);
                                        DataTable dataTable3 = dataSetX_3.Tables[configs.NY22.SheetName];

                                        for (int i = 21; i < dataTable3.Rows.Count - 29; i++)
                                        {
                                            Top10_CPNYGTMAX_HNX cpnygtmax = new Top10_CPNYGTMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpnygtmax.Symbol = dataTable3.Rows[i][column_3[0]].ToString();
                                            if (float.TryParse(dataTable3.Rows[i][column_3[1]].ToString(), out view))
                                            {
                                                cpnygtmax.ClosePrice = Convert.ToDouble(dataTable3.Rows[i][column_3[1]]);
                                            }
                                            else { cpnygtmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[2]].ToString(), out view))
                                            {

                                                cpnygtmax.KLGD = Convert.ToDouble(dataTable3.Rows[i][column_3[2]]);

                                            }
                                            else { cpnygtmax.KLGD = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[3]].ToString(), out view))
                                            {

                                                cpnygtmax.GTNY = Convert.ToDouble(dataTable3.Rows[i][column_3[3]]);

                                            }
                                            else { cpnygtmax.GTNY = 0; }

                                            cpnygtmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, cpnygtmax, null, null, null, null, null, null, null);
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPNYGTMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_4 = configTable.DatTenNY22_4(dataSet);
                                        DataTable dataTable4 = dataSetX_4.Tables[configs.NY22.SheetName];
                                        for (int i = 35; i < dataTable4.Rows.Count - 14; i++)
                                        {
                                            Top10_CPMUAMAX_HNX cpmuamax = new Top10_CPMUAMAX_HNX();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpmuamax.Symbol = dataTable4.Rows[i][column_4[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_4[1]].ToString(), out view))
                                            {
                                                cpmuamax.KLGD = Convert.ToDouble(dataTable4.Rows[i][column_4[1]]);
                                            }
                                            else { cpmuamax.KLGD = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[2]].ToString(), out view))
                                            {

                                                cpmuamax.GTMUA = Convert.ToDouble(dataTable4.Rows[i][column_4[2]]);

                                            }
                                            else { cpmuamax.GTMUA = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[3]].ToString(), out view))
                                            {

                                                cpmuamax.KLNG = Convert.ToDouble(dataTable4.Rows[i][column_4[3]]);

                                            }
                                            else { cpmuamax.KLNG = 0; }

                                            cpmuamax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, cpmuamax, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPMUAMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_5 = configTable.DatTenNY22_5(dataSet);
                                        DataTable dataTable5 = dataSetX_5.Tables[configs.NY22.SheetName];

                                        for (int i = 49; i < dataTable5.Rows.Count - 0; i++)
                                        {
                                            Top10_CPTANGPRICE_HNX cptangprice = new Top10_CPTANGPRICE_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cptangprice.Symbol = dataTable5.Rows[i][column_5[0]].ToString();
                                            if (float.TryParse(dataTable5.Rows[i][column_5[1]].ToString(), out view))
                                            {
                                                cptangprice.ClosePrice = Convert.ToDouble(dataTable5.Rows[i][column_5[1]]);
                                            }
                                            else { cptangprice.ClosePrice = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[2]].ToString(), out view))
                                            {

                                                cptangprice.MucTang = Convert.ToDouble(dataTable5.Rows[i][column_5[2]]);

                                            }
                                            else { cptangprice.MucTang = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[3]].ToString(), out view))
                                            {

                                                cptangprice.TyLeTang = Convert.ToDouble(dataTable5.Rows[i][column_5[3]]);

                                            }
                                            else { cptangprice.TyLeTang = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[4]].ToString(), out view))
                                            {

                                                cptangprice.KLGD = Convert.ToDouble(dataTable5.Rows[i][column_5[4]]);

                                            }
                                            else { cptangprice.KLGD = 0; }

                                            cptangprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, cptangprice, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPTANGPRICE_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        var dataSetX_6 = configTable.DatTenNY22_6(dataSet);
                                        DataTable dataTable6 = dataSetX_6.Tables[configs.NY22.SheetName];

                                        for (int i = 5; i < dataTable6.Rows.Count - 44; i++)
                                        {
                                            Top10_KLGDMAX_HNX klgdmax = new Top10_KLGDMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            klgdmax.Symbol = dataTable6.Rows[i][column_6[0]].ToString();
                                            if (float.TryParse(dataTable6.Rows[i][column_6[1]].ToString(), out view))
                                            {
                                                klgdmax.ClosePrice = Convert.ToDouble(dataTable6.Rows[i][column_6[1]]);
                                            }
                                            else { klgdmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable6.Rows[i][column_6[2]].ToString(), out view))
                                            {

                                                klgdmax.KLGD = Convert.ToDouble(dataTable6.Rows[i][column_6[2]]);

                                            }
                                            else { klgdmax.KLGD = 0; }
                                            if (float.TryParse(dataTable6.Rows[i][column_6[3]].ToString(), out view))
                                            {

                                                klgdmax.TyTrong = Convert.ToDouble(dataTable6.Rows[i][column_6[3]]);

                                            }
                                            else { klgdmax.TyTrong = 0; }


                                            klgdmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, klgdmax, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_KLGDMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        var dataSetX_7 = configTable.DatTenNY22_7(dataSet);
                                        DataTable dataTable7 = dataSetX_7.Tables[configs.NY22.SheetName];

                                        for (int i = 21; i < dataTable7.Rows.Count - 28; i++)
                                        {
                                            Top10_CPGTVHMAX_HNX cpgtvhmax = new Top10_CPGTVHMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgtvhmax.Symbol = dataTable7.Rows[i][column_7[0]].ToString();
                                            if (float.TryParse(dataTable7.Rows[i][column_7[1]].ToString(), out view))
                                            {
                                                cpgtvhmax.ClosePrice = Convert.ToDouble(dataTable7.Rows[i][column_7[1]]);
                                            }
                                            else { cpgtvhmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable7.Rows[i][column_7[2]].ToString(), out view))
                                            {

                                                cpgtvhmax.KLGD = Convert.ToDouble(dataTable7.Rows[i][column_7[2]]);

                                            }
                                            else { cpgtvhmax.KLGD = 0; }
                                            if (float.TryParse(dataTable7.Rows[i][column_7[3]].ToString(), out view))
                                            {

                                                cpgtvhmax.GTVHTT = Convert.ToDouble(dataTable7.Rows[i][column_7[3]]);

                                            }
                                            else { cpgtvhmax.GTVHTT = 0; }


                                            cpgtvhmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, cpgtvhmax, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPGTVHMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        var dataSetX_8 = configTable.DatTenNY22_8(dataSet);
                                        DataTable dataTable8 = dataSetX_8.Tables[configs.NY22.SheetName];

                                        for (int i = 35; i < dataTable8.Rows.Count - 14; i++)
                                        {
                                            Top10_CPBANMAX_HNX cpbanmax = new Top10_CPBANMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpbanmax.Symbol = dataTable8.Rows[i][column_8[0]].ToString();
                                            if (float.TryParse(dataTable8.Rows[i][column_8[1]].ToString(), out view))
                                            {
                                                cpbanmax.KLBAN = Convert.ToDouble(dataTable8.Rows[i][column_8[1]]);
                                            }
                                            else { cpbanmax.KLBAN = 0; }
                                            if (float.TryParse(dataTable8.Rows[i][column_8[2]].ToString(), out view))
                                            {

                                                cpbanmax.GTBAN = Convert.ToDouble(dataTable8.Rows[i][column_8[2]]);

                                            }
                                            else { cpbanmax.GTBAN = 0; }
                                            if (float.TryParse(dataTable8.Rows[i][column_8[3]].ToString(), out view))
                                            {

                                                cpbanmax.KLNG = Convert.ToDouble(dataTable8.Rows[i][column_8[3]]);

                                            }
                                            else { cpbanmax.KLNG = 0; }


                                            cpbanmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, cpbanmax, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPBANMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        var dataSetX_9 = configTable.DatTenNY22_9(dataSet);
                                        DataTable dataTable9 = dataSetX_9.Tables[configs.NY22.SheetName];

                                        for (int i = 49; i < dataTable9.Rows.Count - 0; i++)
                                        {
                                            Top10_CPGIAMPRICE_HNX cpgiamprice = new Top10_CPGIAMPRICE_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgiamprice.Symbol = dataTable9.Rows[i][column_9[0]].ToString();
                                            if (float.TryParse(dataTable9.Rows[i][column_9[1]].ToString(), out view))
                                            {
                                                cpgiamprice.ClosePrice = Convert.ToDouble(dataTable9.Rows[i][column_9[1]]);
                                            }
                                            else { cpgiamprice.ClosePrice = 0; }
                                            if (float.TryParse(dataTable9.Rows[i][column_9[2]].ToString(), out view))
                                            {

                                                cpgiamprice.MucGiam = Convert.ToDouble(dataTable9.Rows[i][column_9[2]]);

                                            }
                                            else { cpgiamprice.MucGiam = 0; }
                                            if (float.TryParse(dataTable9.Rows[i][column_9[3]].ToString(), out view))
                                            {

                                                cpgiamprice.TyLeGiam = Convert.ToDouble(dataTable9.Rows[i][column_9[3]]);

                                            }
                                            else { cpgiamprice.TyLeGiam = 0; }

                                            if (float.TryParse(dataTable9.Rows[i][column_9[4]].ToString(), out view))
                                            {

                                                cpgiamprice.KLGD = Convert.ToDouble(dataTable9.Rows[i][column_9[4]]);

                                            }
                                            else { cpgiamprice.TyLeGiam = 0; }

                                            cpgiamprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, null, cpgiamprice, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }


                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY22.Data_Table_Top10_CPGIAMPRICE_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        default:
                            // code block
                            Console.WriteLine("File chưa định dạng để read!");
                            break;
                    }
                    // Console.WriteLine("Done!");
                }
                if (match_3.Success)
                {
                    string dateString = "";
                    dateString = match_3.Groups[1].Value;
                    string thitruong = match_3.Groups[2].Value;
                    if (dateString == "01304021")
                    {
                        dateString = "20130402";
                    }
                    DateTime date = DateTime.ParseExact(dateString, "yyyyMMdd", CultureInfo.InvariantCulture);
                    string dateS = date.ToString("yyyy-MM-dd");
                    DateTime dateFile = DateTime.ParseExact(dateS, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    DateTime dateTo = DateTime.ParseExact(configs.ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    switch (thitruong)
                    {
                        case ConfigApp.NY_4:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY_GDNDTNN_Phien(dataSet);
                                        float view;
                                        DataTable dataTable;
                                        if (dateString == "20130627")
                                        {
                                            dataTable = dataSetX.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable = dataSetX.Tables[configs.NY_GDNDTNN_Phien.SheetName];

                                        }

                                        string[] column = configs.NY_GDNDTNN_Phien.BeginCell.Split(',');
                                        for (int i = 5; i < dataTable.Rows.Count - 26; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                GIAODICHNHADAUTUNN_HNX gdndtnn_hnx = new GIAODICHNHADAUTUNN_HNX();

                                                gdndtnn_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                gdndtnn_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLMUA_KL = 0;

                                                }
                                                else { gdndtnn_hnx.KLMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTMUA_KL = 0;

                                                }
                                                else { gdndtnn_hnx.GTMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLBAN_KL = 0;

                                                }
                                                else { gdndtnn_hnx.KLBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTBAN_KL = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn_hnx.GTBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLMUA_TT = 0;

                                                }
                                                else { gdndtnn_hnx.KLMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTMUA_TT = 0;

                                                }
                                                else { gdndtnn_hnx.GTMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLBAN_TT = 0;

                                                }
                                                else { gdndtnn_hnx.KLBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTBAN_TT = 0;

                                                }
                                                else { gdndtnn_hnx.GTBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLMUA_TC = 0;

                                                }
                                                else { gdndtnn_hnx.KLMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTMUA_TC = 0;

                                                }
                                                else
                                                {
                                                    gdndtnn_hnx.GTMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLBAN_TC = 0;

                                                }
                                                else { gdndtnn_hnx.KLBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.GTBAN_TC = 0;

                                                }
                                                else { gdndtnn_hnx.GTBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLCK_MAX = 0;

                                                }
                                                else { gdndtnn_hnx.KLCK_MAX = Convert.ToDouble(dataTable.Rows[i][column[14]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLCK_NDTNN = 0;

                                                }
                                                else { gdndtnn_hnx.KLCK_NDTNN = Convert.ToDouble(dataTable.Rows[i][column[15]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    gdndtnn_hnx.KLCK_CDPNG = 0;

                                                }
                                                else { gdndtnn_hnx.KLCK_CDPNG = Convert.ToDouble(dataTable.Rows[i][column[16]]); }
                                                gdndtnn_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX(null, gdndtnn_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_GDNDTNN_Phien.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_1:
                            try
                            {
                                if (dateFile < dateTo)
                                {
                                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                    {
                                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                                        {
                                            EBulkScript eBulkScript = new EBulkScript();
                                            var dataSet = reader.AsDataSet();

                                            int count = dataSet.Tables[0].Columns.Count;
                                            if (count == 21)
                                            {
                                                var dataSetX = configTable.DatTenNY_KQGD_Phien2SS(dataSet);
                                                float view;
                                                DataTable dataTable;
                                                if (dateString == "20130627")
                                                {
                                                    dataTable = dataSetX.Tables["Sheet1"];
                                                }
                                                else
                                                {
                                                    dataTable = dataSetX.Tables[configs.NY_KQGD_Phien2.SheetName];

                                                }
                                                string[] column = configs.NY_KQGD_Phien2.BeginCell.Split(',');
                                                for (int i = 6; i < dataTable.Rows.Count - 1; i++)
                                                {
                                                    if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                    {
                                                        KQGIAODICHCP_HNX_2013_2 kq_hnx2 = new KQGIAODICHCP_HNX_2013_2();


                                                        kq_hnx2.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                        kq_hnx2.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                        if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                        {

                                                            kq_hnx2.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                        }
                                                        else { kq_hnx2.BasicPrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                        {

                                                            kq_hnx2.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                        }
                                                        else { kq_hnx2.OpenPrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                        {

                                                            kq_hnx2.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                        }
                                                        else { kq_hnx2.ClosePrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                        {

                                                            kq_hnx2.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                        }
                                                        else { kq_hnx2.HighestPrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                        {

                                                            kq_hnx2.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                        }
                                                        else { kq_hnx2.LowestPrice = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GiaCoSo = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                        }
                                                        else { kq_hnx2.GiaCoSo = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                        }
                                                        else { kq_hnx2.TDDiem = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                        }
                                                        else { kq_hnx2.TDPhanTram = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLGD_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                        }
                                                        else { kq_hnx2.KLGD_KL = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTGD_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                        }
                                                        else { kq_hnx2.GTGD_KL = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLGD_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                        }
                                                        else { kq_hnx2.KLGD_TT = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTGD_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                        }
                                                        else { kq_hnx2.GTGD_TT = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                        }
                                                        else { kq_hnx2.KLGD_TC = 0; }



                                                        if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                        }
                                                        else { kq_hnx2.TITRONG1 = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                        }
                                                        else { kq_hnx2.GTGD_TC = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                        }
                                                        else { kq_hnx2.TITRONG2 = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                        }
                                                        else { kq_hnx2.KLCPLH = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                        }
                                                        else { kq_hnx2.GTVHTT_GT = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                        }
                                                        else { kq_hnx2.GTVHTT_TT = 0; }

                                                        kq_hnx2.VonDL = 0;

                                                        kq_hnx2.Trangding_Date = dateFile;
                                                        eBulkScript = this.configTable.GetScriptTTCBHNX_2013(null, null, kq_hnx2);
                                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                    }
                                                }


                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_KQGD_Phien2.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }

                                            }
                                            else
                                            {
                                                var dataSetX = configTable.DatTenNY_KQGD_Phien2(dataSet);
                                                float view;
                                                DataTable dataTable;
                                                if (dateString == "20130627")
                                                {
                                                    dataTable = dataSetX.Tables["Sheet1"];
                                                }
                                                else
                                                {
                                                    dataTable = dataSetX.Tables[configs.NY_KQGD_Phien2.SheetName];

                                                }
                                                string[] column = configs.NY_KQGD_Phien2.BeginCell.Split(',');
                                                for (int i = 6; i < dataTable.Rows.Count - 1; i++)
                                                {
                                                    if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                    {
                                                        KQGIAODICHCP_HNX_2013_2 kq_hnx2 = new KQGIAODICHCP_HNX_2013_2();


                                                        kq_hnx2.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                        kq_hnx2.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                        if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                        {

                                                            kq_hnx2.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                        }
                                                        else { kq_hnx2.BasicPrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                        {

                                                            kq_hnx2.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                        }
                                                        else { kq_hnx2.OpenPrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                        {

                                                            kq_hnx2.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                        }
                                                        else { kq_hnx2.ClosePrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                        {

                                                            kq_hnx2.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                        }
                                                        else { kq_hnx2.HighestPrice = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                        {

                                                            kq_hnx2.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                        }
                                                        else { kq_hnx2.LowestPrice = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GiaCoSo = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                        }
                                                        else { kq_hnx2.GiaCoSo = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                        }
                                                        else { kq_hnx2.TDDiem = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                        }
                                                        else { kq_hnx2.TDPhanTram = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLGD_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                        }
                                                        else { kq_hnx2.KLGD_KL = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTGD_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                        }
                                                        else { kq_hnx2.GTGD_KL = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLGD_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                        }
                                                        else { kq_hnx2.KLGD_TT = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTGD_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                        }
                                                        else { kq_hnx2.GTGD_TT = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                        }
                                                        else { kq_hnx2.KLGD_TC = 0; }



                                                        if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                        }
                                                        else { kq_hnx2.TITRONG1 = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                        }
                                                        else { kq_hnx2.GTGD_TC = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                        {

                                                            kq_hnx2.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                        }
                                                        else { kq_hnx2.TITRONG2 = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                        {

                                                            kq_hnx2.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                        }
                                                        else { kq_hnx2.KLCPLH = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                        }
                                                        else { kq_hnx2.GTVHTT_GT = 0; }
                                                        if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                        {

                                                            kq_hnx2.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                        }
                                                        else { kq_hnx2.GTVHTT_TT = 0; }

                                                        if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                        {

                                                            kq_hnx2.VonDL = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                                        }
                                                        else { kq_hnx2.VonDL = 0; }

                                                        kq_hnx2.Trangding_Date = dateFile;
                                                        eBulkScript = this.configTable.GetScriptTTCBHNX_2013(null, null, kq_hnx2);
                                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                    }
                                                }


                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_KQGD_Phien2.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }

                                            }


                                        }
                                    }
                                }
                                else
                                {
                                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                    {
                                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                                        {
                                            EBulkScript eBulkScript = new EBulkScript();
                                            var dataSet = reader.AsDataSet();


                                            var dataSetX = configTable.DatTenNY_KQGD_Phien(dataSet);
                                            float view;

                                            DataTable dataTable;
                                            if (dateString == "20130627")
                                            {
                                                dataTable = dataSetX.Tables["Sheet1"];
                                            }
                                            else
                                            {
                                                dataTable = dataSetX.Tables[configs.NY_KQGD_Phien.SheetName];

                                            }
                                            string[] column = configs.NY_KQGD_Phien.BeginCell.Split(',');
                                            for (int i = 6; i < dataTable.Rows.Count - 20; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {
                                                    KQGIAODICHCP_HNX_2013 kq_hnx = new KQGIAODICHCP_HNX_2013();


                                                    kq_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    kq_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {

                                                        kq_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                    }
                                                    else { kq_hnx.BasicPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {

                                                        kq_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                    }
                                                    else { kq_hnx.OpenPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {

                                                        kq_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                    }
                                                    else { kq_hnx.ClosePrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {

                                                        kq_hnx.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                    }
                                                    else { kq_hnx.HighestPrice = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {

                                                        kq_hnx.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                    }
                                                    else { kq_hnx.LowestPrice = 0; }


                                                    if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {

                                                        kq_hnx.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                    }
                                                    else { kq_hnx.TDDiem = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {

                                                        kq_hnx.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                    }
                                                    else { kq_hnx.TDPhanTram = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGD_KL = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                    }
                                                    else { kq_hnx.KLGD_KL = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGD_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                    }
                                                    else { kq_hnx.GTGD_KL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGD_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                    }
                                                    else { kq_hnx.KLGD_TT = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGD_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                    }
                                                    else { kq_hnx.GTGD_TT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGD_LL = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                    }
                                                    else { kq_hnx.KLGD_LL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGD_LL = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                    }
                                                    else { kq_hnx.GTGD_LL = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                    }
                                                    else { kq_hnx.KLGD_TC = 0; }



                                                    if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                    {

                                                        kq_hnx.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                    }
                                                    else { kq_hnx.TITRONG1 = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                    }
                                                    else { kq_hnx.GTGD_TC = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                    {

                                                        kq_hnx.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                    }
                                                    else { kq_hnx.TITRONG2 = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                    {

                                                        kq_hnx.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                    }
                                                    else { kq_hnx.KLCPLH = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                    }
                                                    else { kq_hnx.GTVHTT_GT = 0; }
                                                    if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                                    {

                                                        kq_hnx.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                                    }
                                                    else { kq_hnx.GTVHTT_TT = 0; }

                                                    if (float.TryParse(dataTable.Rows[i][column[22]].ToString(), out view))
                                                    {

                                                        kq_hnx.VonDL = Convert.ToDouble(dataTable.Rows[i][column[22]]);

                                                    }
                                                    else { kq_hnx.VonDL = 0; }

                                                    kq_hnx.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX_2013(kq_hnx, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                                }
                                            }


                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_KQGD_Phien.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }


                                        }
                                    }

                                }
                            }

                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_3:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY_EOD4_View(dataSet);
                                        float view;

                                        DataTable dataTable;
                                        if (dateString == "20130627")
                                        {
                                            dataTable = dataSetX.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable = dataSetX.Tables[configs.NY_EOD4_View.SheetName];

                                        }
                                        string[] column = configs.NY_EOD4_View.BeginCell.Split(',');
                                        for (int i = 5; i < dataTable.Rows.Count - 19; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {

                                                TKCUNGCAUTTCP_HNX cc_hnx = new TKCUNGCAUTTCP_HNX();


                                                cc_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc_hnx.SLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc_hnx.KLDATMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc_hnx.SLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc_hnx.KLDATBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc_hnx.SLDATMUA_TT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc_hnx.KLDATMUA_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc_hnx.SLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc_hnx.KLDATBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc_hnx.SLDATMUA_TC = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc_hnx.KLDATMUA_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {

                                                    cc_hnx.SLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                }
                                                else { cc_hnx.SLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDATBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                }
                                                else { cc_hnx.KLDATBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDUMUA = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                }
                                                else { cc_hnx.KLDUMUA = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {

                                                    cc_hnx.KLDUBAN = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                }
                                                else { cc_hnx.KLDUBAN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {

                                                    cc_hnx.KLTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                }
                                                else { cc_hnx.KLTHUCHIEN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                {

                                                    cc_hnx.GTTHUCHIEN = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                }
                                                else { cc_hnx.GTTHUCHIEN = 0; }
                                                cc_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, cc_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_EOD4_View.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_2:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX_1 = configTable.DatTenNY22_1(dataSet);
                                        var dataSetX_2 = configTable.DatTenNY22_2(dataSet);

                                        float view;
                                        //  DataTable dataTable;

                                        DataTable dataTable;
                                        DataTable dataTable2;

                                        if (dateString == "20130627")
                                        {
                                            dataTable = dataSetX_1.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable = dataSetX_1.Tables[configs.NY_TT_Phien.SheetName];

                                        }
                                        if (dateString == "20130627")
                                        {
                                            dataTable2 = dataSetX_2.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable2 = dataSetX_2.Tables[configs.NY_TT_Phien.SheetName];

                                        }
                                        string[] column = configs.NY_TT_Phien.Data_Table_Chi_Tieu_HNX.BeginCell.Split(',');
                                        string[] column_2 = configs.NY_TT_Phien.Data_Table_Top10_CPGDMAX_HNX.BeginCell.Split(',');
                                        string[] column_3 = configs.NY_TT_Phien.Data_Table_Top10_CPNYGTMAX_HNX.BeginCell.Split(',');
                                        string[] column_4 = configs.NY_TT_Phien.Data_Table_Top10_CPMUAMAX_HNX.BeginCell.Split(',');
                                        string[] column_5 = configs.NY_TT_Phien.Data_Table_Top10_CPTANGPRICE_HNX.BeginCell.Split(',');

                                        string[] column_6 = configs.NY_TT_Phien.Data_Table_Top10_KLGDMAX_HNX.BeginCell.Split(',');
                                        string[] column_7 = configs.NY_TT_Phien.Data_Table_Top10_CPGTVHMAX_HNX.BeginCell.Split(',');
                                        string[] column_8 = configs.NY_TT_Phien.Data_Table_Top10_CPBANMAX_HNX.BeginCell.Split(',');
                                        string[] column_9 = configs.NY_TT_Phien.Data_Table_Top10_CPGIAMPRICE_HNX.BeginCell.Split(',');

                                        for (int i = 4; i < dataTable.Rows.Count - 32; i++)
                                        {
                                            Chi_Tieu_HNX ct = new Chi_Tieu_HNX();

                                            //Chi_Tieu,Don_Vi,So_Lieu,Trangding_Date

                                            ct.Chi_Tieu = dataTable.Rows[i][column[0]].ToString();
                                            // ct.Don_Vi = dataTable.Rows[i][column[1]].ToString();
                                            if (float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {

                                                ct.So_Lieu = Convert.ToDouble(dataTable.Rows[i][column[1]]);

                                            }
                                            else { ct.So_Lieu = 0; }

                                            ct.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, ct, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Chi_Tieu_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        for (int i = 5; i < dataTable2.Rows.Count - 44; i++)
                                        {
                                            Top10_CPGDMAX_HNX cpgdmax = new Top10_CPGDMAX_HNX();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpgdmax.Symbol = dataTable2.Rows[i][column_2[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_2[1]].ToString(), out view))
                                            {
                                                cpgdmax.ClosePrice = Convert.ToDouble(dataTable2.Rows[i][column_2[1]]);
                                            }
                                            else { cpgdmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[2]].ToString(), out view))
                                            {
                                                cpgdmax.GTGD = Convert.ToDouble(dataTable2.Rows[i][column_2[2]]);
                                            }
                                            else { cpgdmax.GTGD = 0; }
                                            if (float.TryParse(dataTable2.Rows[i][column_2[3]].ToString(), out view))
                                            {

                                                cpgdmax.TyTrong = Convert.ToDouble(dataTable2.Rows[i][column_2[3]]);

                                            }
                                            else { cpgdmax.TyTrong = 0; }

                                            cpgdmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, cpgdmax, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPGDMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_3 = configTable.DatTenNY22_3(dataSet);
                                        DataTable dataTable3;
                                        //= dataSetX_3.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable3 = dataSetX_3.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable3 = dataSetX_3.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 21; i < dataTable3.Rows.Count - 29; i++)
                                        {
                                            Top10_CPNYGTMAX_HNX cpnygtmax = new Top10_CPNYGTMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpnygtmax.Symbol = dataTable3.Rows[i][column_3[0]].ToString();
                                            if (float.TryParse(dataTable3.Rows[i][column_3[1]].ToString(), out view))
                                            {
                                                cpnygtmax.ClosePrice = Convert.ToDouble(dataTable3.Rows[i][column_3[1]]);
                                            }
                                            else { cpnygtmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[2]].ToString(), out view))
                                            {

                                                cpnygtmax.KLGD = Convert.ToDouble(dataTable3.Rows[i][column_3[2]]);

                                            }
                                            else { cpnygtmax.KLGD = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[3]].ToString(), out view))
                                            {

                                                cpnygtmax.GTNY = Convert.ToDouble(dataTable3.Rows[i][column_3[3]]);

                                            }
                                            else { cpnygtmax.GTNY = 0; }

                                            cpnygtmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, cpnygtmax, null, null, null, null, null, null, null);
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPNYGTMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_4 = configTable.DatTenNY22_4(dataSet);
                                        DataTable dataTable4;
                                        //= dataSetX_4.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable4 = dataSetX_4.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable4 = dataSetX_4.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 35; i < dataTable4.Rows.Count - 14; i++)
                                        {
                                            Top10_CPMUAMAX_HNX cpmuamax = new Top10_CPMUAMAX_HNX();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            cpmuamax.Symbol = dataTable4.Rows[i][column_4[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_4[1]].ToString(), out view))
                                            {
                                                cpmuamax.KLGD = Convert.ToDouble(dataTable4.Rows[i][column_4[1]]);
                                            }
                                            else { cpmuamax.KLGD = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[2]].ToString(), out view))
                                            {

                                                cpmuamax.GTMUA = Convert.ToDouble(dataTable4.Rows[i][column_4[2]]);

                                            }
                                            else { cpmuamax.GTMUA = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[3]].ToString(), out view))
                                            {

                                                cpmuamax.KLNG = Convert.ToDouble(dataTable4.Rows[i][column_4[3]]);

                                            }
                                            else { cpmuamax.KLNG = 0; }

                                            cpmuamax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, cpmuamax, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPMUAMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_5 = configTable.DatTenNY22_5(dataSet);
                                        DataTable dataTable5;
                                        //= dataSetX_5.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable5 = dataSetX_5.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable5 = dataSetX_5.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 49; i < dataTable5.Rows.Count - 0; i++)
                                        {
                                            Top10_CPTANGPRICE_HNX cptangprice = new Top10_CPTANGPRICE_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cptangprice.Symbol = dataTable5.Rows[i][column_5[0]].ToString();
                                            if (float.TryParse(dataTable5.Rows[i][column_5[1]].ToString(), out view))
                                            {
                                                cptangprice.ClosePrice = Convert.ToDouble(dataTable5.Rows[i][column_5[1]]);
                                            }
                                            else { cptangprice.ClosePrice = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[2]].ToString(), out view))
                                            {

                                                cptangprice.MucTang = Convert.ToDouble(dataTable5.Rows[i][column_5[2]]);

                                            }
                                            else { cptangprice.MucTang = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[3]].ToString(), out view))
                                            {

                                                cptangprice.TyLeTang = Convert.ToDouble(dataTable5.Rows[i][column_5[3]]);

                                            }
                                            else { cptangprice.TyLeTang = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[4]].ToString(), out view))
                                            {

                                                cptangprice.KLGD = Convert.ToDouble(dataTable5.Rows[i][column_5[4]]);

                                            }
                                            else { cptangprice.KLGD = 0; }

                                            cptangprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, cptangprice, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPTANGPRICE_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_6 = configTable.DatTenNY22_6(dataSet);
                                        DataTable dataTable6;
                                        //= dataSetX_6.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable6 = dataSetX_6.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable6 = dataSetX_6.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 5; i < dataTable6.Rows.Count - 44; i++)
                                        {
                                            Top10_KLGDMAX_HNX klgdmax = new Top10_KLGDMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            klgdmax.Symbol = dataTable6.Rows[i][column_6[0]].ToString();
                                            if (float.TryParse(dataTable6.Rows[i][column_6[1]].ToString(), out view))
                                            {
                                                klgdmax.ClosePrice = Convert.ToDouble(dataTable6.Rows[i][column_6[1]]);
                                            }
                                            else { klgdmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable6.Rows[i][column_6[2]].ToString(), out view))
                                            {

                                                klgdmax.KLGD = Convert.ToDouble(dataTable6.Rows[i][column_6[2]]);

                                            }
                                            else { klgdmax.KLGD = 0; }
                                            if (float.TryParse(dataTable6.Rows[i][column_6[3]].ToString(), out view))
                                            {

                                                klgdmax.TyTrong = Convert.ToDouble(dataTable6.Rows[i][column_6[3]]);

                                            }
                                            else { klgdmax.TyTrong = 0; }


                                            klgdmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, klgdmax, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_KLGDMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_7 = configTable.DatTenNY22_7(dataSet);
                                        DataTable dataTable7;
                                        //= dataSetX_7.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable7 = dataSetX_7.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable7 = dataSetX_7.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 21; i < dataTable7.Rows.Count - 28; i++)
                                        {
                                            Top10_CPGTVHMAX_HNX cpgtvhmax = new Top10_CPGTVHMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgtvhmax.Symbol = dataTable7.Rows[i][column_7[0]].ToString();
                                            if (float.TryParse(dataTable7.Rows[i][column_7[1]].ToString(), out view))
                                            {
                                                cpgtvhmax.ClosePrice = Convert.ToDouble(dataTable7.Rows[i][column_7[1]]);
                                            }
                                            else { cpgtvhmax.ClosePrice = 0; }
                                            if (float.TryParse(dataTable7.Rows[i][column_7[2]].ToString(), out view))
                                            {

                                                cpgtvhmax.KLGD = Convert.ToDouble(dataTable7.Rows[i][column_7[2]]);

                                            }
                                            else { cpgtvhmax.KLGD = 0; }
                                            if (float.TryParse(dataTable7.Rows[i][column_7[3]].ToString(), out view))
                                            {

                                                cpgtvhmax.GTVHTT = Convert.ToDouble(dataTable7.Rows[i][column_7[3]]);

                                            }
                                            else { cpgtvhmax.GTVHTT = 0; }


                                            cpgtvhmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, cpgtvhmax, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPGTVHMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_8 = configTable.DatTenNY22_8(dataSet);
                                        DataTable dataTable8;
                                        //= dataSetX_8.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable8 = dataSetX_8.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable8 = dataSetX_8.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 35; i < dataTable8.Rows.Count - 14; i++)
                                        {
                                            Top10_CPBANMAX_HNX cpbanmax = new Top10_CPBANMAX_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpbanmax.Symbol = dataTable8.Rows[i][column_8[0]].ToString();
                                            if (float.TryParse(dataTable8.Rows[i][column_8[1]].ToString(), out view))
                                            {
                                                cpbanmax.KLBAN = Convert.ToDouble(dataTable8.Rows[i][column_8[1]]);
                                            }
                                            else { cpbanmax.KLBAN = 0; }
                                            if (float.TryParse(dataTable8.Rows[i][column_8[2]].ToString(), out view))
                                            {

                                                cpbanmax.GTBAN = Convert.ToDouble(dataTable8.Rows[i][column_8[2]]);

                                            }
                                            else { cpbanmax.GTBAN = 0; }
                                            if (float.TryParse(dataTable8.Rows[i][column_8[3]].ToString(), out view))
                                            {

                                                cpbanmax.KLNG = Convert.ToDouble(dataTable8.Rows[i][column_8[3]]);

                                            }
                                            else { cpbanmax.KLNG = 0; }


                                            cpbanmax.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, cpbanmax, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPBANMAX_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        var dataSetX_9 = configTable.DatTenNY22_9(dataSet);
                                        DataTable dataTable9;
                                        //= dataSetX_9.Tables[configs.NY22.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable9 = dataSetX_9.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable9 = dataSetX_9.Tables[configs.NY22.SheetName];

                                        }
                                        for (int i = 49; i < dataTable9.Rows.Count - 0; i++)
                                        {
                                            Top10_CPGIAMPRICE_HNX cpgiamprice = new Top10_CPGIAMPRICE_HNX();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgiamprice.Symbol = dataTable9.Rows[i][column_9[0]].ToString();
                                            if (float.TryParse(dataTable9.Rows[i][column_9[1]].ToString(), out view))
                                            {
                                                cpgiamprice.ClosePrice = Convert.ToDouble(dataTable9.Rows[i][column_9[1]]);
                                            }
                                            else { cpgiamprice.ClosePrice = 0; }
                                            if (float.TryParse(dataTable9.Rows[i][column_9[2]].ToString(), out view))
                                            {

                                                cpgiamprice.MucGiam = Convert.ToDouble(dataTable9.Rows[i][column_9[2]]);

                                            }
                                            else { cpgiamprice.MucGiam = 0; }
                                            if (float.TryParse(dataTable9.Rows[i][column_9[3]].ToString(), out view))
                                            {

                                                cpgiamprice.TyLeGiam = Convert.ToDouble(dataTable9.Rows[i][column_9[3]]);

                                            }
                                            else { cpgiamprice.TyLeGiam = 0; }

                                            if (float.TryParse(dataTable9.Rows[i][column_9[4]].ToString(), out view))
                                            {

                                                cpgiamprice.KLGD = Convert.ToDouble(dataTable9.Rows[i][column_9[4]]);

                                            }
                                            else { cpgiamprice.TyLeGiam = 0; }

                                            cpgiamprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX(null, null, null, null, null, null, null, null, null, null, null, null, null, cpgiamprice, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }


                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_TT_Phien.Data_Table_Top10_CPGIAMPRICE_HNX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_5:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenNY25(dataSet);
                                        float view;
                                        DataTable dataTable;
                                        //= dataSetX.Tables[configs.NY_CPNY_Phien.SheetName];
                                        if (dateString == "20130627")
                                        {
                                            dataTable = dataSetX.Tables["Sheet1"];
                                        }
                                        else
                                        {
                                            dataTable = dataSetX.Tables[configs.NY_CPNY_Phien.SheetName];

                                        }
                                        //  Console.WriteLine(filePath );
                                        string[] column = configs.NY_CPNY_Phien.BeginCell.Split(',');
                                        for (int i = 4; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NY_TTCP_2013 ttcb = new NY_TTCP_2013();

                                                ttcb.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);

                                                ttcb.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2014,Co_Tuc_2015,PE,EPS2015,ROE2015,

                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    ttcb.KLCP_NY = 0;

                                                }
                                                else { ttcb.KLCP_NY = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    ttcb.KLCP_LH = 0;

                                                }
                                                else { ttcb.KLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    ttcb.Co_Tuc_2013 = 0;

                                                }
                                                else { ttcb.Co_Tuc_2013 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    ttcb.Co_Tuc_2014 = 0;

                                                }
                                                else
                                                {
                                                    ttcb.Co_Tuc_2014 = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    ttcb.PE = 0;

                                                }
                                                else { ttcb.PE = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    ttcb.EPS2014 = 0;

                                                }
                                                else { ttcb.EPS2014 = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    ttcb.ROE2014 = 0;

                                                }
                                                else { ttcb.ROE2014 = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    ttcb.ROA2014 = 0;

                                                }
                                                else { ttcb.ROA2014 = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    ttcb.BasicPrice_KT = 0;

                                                }
                                                else { ttcb.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    ttcb.CeilingPrice_KT = 0;

                                                }
                                                else { ttcb.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    ttcb.FloorPrice_KT = 0;

                                                }
                                                else { ttcb.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                //ROA2015,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
                                                ttcb.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX_2013(null, ttcb, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.NY_CPNY_Phien.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();


                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        default:
                            // code block
                            Console.WriteLine("File chưa định dạng để read!");
                            break;
                    }
                    // Console.WriteLine("Done!");
                }
                if (match_5.Success)
                {
                    string dateString = "";
                    dateString = match_5.Value;
                    DateTime dateTime = DateTime.ParseExact(dateString, "dd-MM-yyyy", CultureInfo.InvariantCulture);
                    string outputDate = dateTime.ToString("yyyy-MM-dd");

                    DateTime dateFile = DateTime.ParseExact(outputDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    // DateTime dateTo = DateTime.ParseExact(configs.ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    if (!filePath.Contains("Bond") && !filePath.Contains("Draft"))
                    {
                        try
                        {


                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];
                                    if (dataSet != null)
                                    {
                                        int counts = dataSet.Columns.Count;
                                        if (counts > 15)
                                        {
                                            if (dataSet.Rows[0][16].ToString().Contains("Cổ tức 2009") && dataSet.Rows[0][15].ToString() == "" && dataSet.Rows[1][14].ToString() == "" && dataSet.Rows[1][13].ToString() == "B/q")
                                            {

                                                DataTable dataTable = configTable.DatTenTT_DKGD_2011(dataSet);


                                                //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                                string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                                for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                                {
                                                    if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                    {

                                                        KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                        dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                        dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_DKGD = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_LH = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2010 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.PE = 0;

                                                        }
                                                        else
                                                        {
                                                            dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                        }
                                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.EPS2010 = 0;

                                                        }
                                                        else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.KLGD_10PHIEN = 0;

                                                        }
                                                        else { dkgd_hnx.KLGD_10PHIEN = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                        //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                        //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROE = 0;

                                                        }
                                                        else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROA = 0;

                                                        }
                                                        else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BasicPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.CeilingPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.FloorPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BinhQuan = 0;

                                                        }
                                                        else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Tong = 0;

                                                        }
                                                        else { dkgd_hnx.Tong = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                        if (!float.TryParse(dataTable.Rows[i]["Column16"].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2009 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2009 = Convert.ToDouble(dataTable.Rows[i]["Column16"]); }
                                                        dkgd_hnx.Trangding_Date = dateFile;
                                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                    }
                                                }
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }
                                                //Console.WriteLine("File: " + filePath);
                                            }

                                            if (dataSet.Rows[0][16].ToString() == "" && dataSet.Rows[0][15].ToString() == "" && dataSet.Rows[1][14].ToString() == "" && dataSet.Rows[1][13].ToString() == "B/q")
                                            {
                                                DataTable dataTable = configTable.DatTenTT_DKGD_2011(dataSet);


                                                //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                                string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                                for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                                {
                                                    if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                    {
                                                        KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                        dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                        dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_DKGD = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_LH = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2010 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.PE = 0;

                                                        }
                                                        else
                                                        {
                                                            dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                        }
                                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.EPS2010 = 0;

                                                        }
                                                        else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.KLGD_10PHIEN = 0;

                                                        }
                                                        else { dkgd_hnx.KLGD_10PHIEN = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                        //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                        //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROE = 0;

                                                        }
                                                        else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROA = 0;

                                                        }
                                                        else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BasicPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.CeilingPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.FloorPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BinhQuan = 0;

                                                        }
                                                        else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Tong = 0;

                                                        }
                                                        else { dkgd_hnx.Tong = Convert.ToDouble(dataTable.Rows[i][column[13]]); }

                                                        dkgd_hnx.Co_Tuc_2009 = 0;

                                                        dkgd_hnx.Trangding_Date = dateFile;
                                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                    }
                                                }
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }
                                                // Console.WriteLine("File2: " + filePath);
                                            }
                                            if (dataSet.Rows[1][15].ToString() == "" && dataSet.Rows[1][14].ToString() == "B/q" && dataSet.Rows[0][13].ToString().Contains("Cổ tức 2009"))
                                            {
                                                DataTable dataTable = configTable.DatTenTT_DKGD_2011(dataSet);


                                                //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                                string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                                for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                                {
                                                    if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                    {
                                                        KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                        dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                        dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_DKGD = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_LH = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2010 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.PE = 0;

                                                        }
                                                        else
                                                        {
                                                            dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                        }
                                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.EPS2010 = 0;

                                                        }
                                                        else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.KLGD_10PHIEN = 0;

                                                        }
                                                        else { dkgd_hnx.KLGD_10PHIEN = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                        //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                        //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROE = 0;

                                                        }
                                                        else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROA = 0;

                                                        }
                                                        else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BasicPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.CeilingPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.FloorPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BinhQuan = 0;

                                                        }
                                                        else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Tong = 0;

                                                        }
                                                        else { dkgd_hnx.Tong = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2009 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2009 = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                        dkgd_hnx.Trangding_Date = dateFile;
                                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                    }
                                                }
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }
                                                //   Console.WriteLine("File3: " + filePath);
                                            }
                                            if (dataSet.Rows[0][15].ToString().Contains("Cổ tức 2009") && dataSet.Rows[1][14].ToString() == "" && dataSet.Rows[1][13].ToString() == "B/q")
                                            {
                                                DataTable dataTable = configTable.DatTenTT_DKGD_2011(dataSet);


                                                //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                                string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                                for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                                {
                                                    if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                    {
                                                        KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                        dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                        dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_DKGD = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.SLCP_LH = 0;

                                                        }
                                                        else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2010 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.PE = 0;

                                                        }
                                                        else
                                                        {
                                                            dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                        }
                                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.EPS2010 = 0;

                                                        }
                                                        else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.KLGD_10PHIEN = 0;

                                                        }
                                                        else { dkgd_hnx.KLGD_10PHIEN = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                        //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                        //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROE = 0;

                                                        }
                                                        else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.ROA = 0;

                                                        }
                                                        else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BasicPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.CeilingPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.FloorPrice_KT = 0;

                                                        }
                                                        else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.BinhQuan = 0;

                                                        }
                                                        else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Tong = 0;

                                                        }
                                                        else { dkgd_hnx.Tong = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                        if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                        {
                                                            dkgd_hnx.Co_Tuc_2009 = 0;

                                                        }
                                                        else { dkgd_hnx.Co_Tuc_2009 = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                        dkgd_hnx.Trangding_Date = dateFile;
                                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                    }
                                                }
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                {
                                                    // exec script mssql+oracle
                                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                    configTable.ExecBulkScript(test);
                                                    mssqlBuilder_HNX.Clear();

                                                }
                                            }

                                        }
                                        else
                                        {
                                            DataTable dataTable = configTable.DatTenTT_DKGD_2011SS(dataSet);


                                            //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                            string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                            for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                            {
                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                                {
                                                    KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                    dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.SLCP_DKGD = 0;

                                                    }
                                                    else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.SLCP_LH = 0;

                                                    }
                                                    else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.Co_Tuc_2010 = 0;

                                                    }
                                                    else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.PE = 0;

                                                    }
                                                    else
                                                    {
                                                        dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                    }
                                                    if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.EPS2010 = 0;

                                                    }
                                                    else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.KLGD_10PHIEN = 0;

                                                    }
                                                    else { dkgd_hnx.KLGD_10PHIEN = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                    //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                    //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                    if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.ROE = 0;

                                                    }
                                                    else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.ROA = 0;

                                                    }
                                                    else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.BasicPrice_KT = 0;

                                                    }
                                                    else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.CeilingPrice_KT = 0;

                                                    }
                                                    else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.FloorPrice_KT = 0;

                                                    }
                                                    else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.BinhQuan = 0;

                                                    }
                                                    else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.Tong = 0;

                                                    }
                                                    else { dkgd_hnx.Tong = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                    {
                                                        dkgd_hnx.Co_Tuc_2009 = 0;

                                                    }
                                                    else { dkgd_hnx.Co_Tuc_2009 = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                    dkgd_hnx.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }
                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                        }
                                    }

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }
                        //Tinh hinh dat lenh
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;

                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenTH_DATLENH_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.TH_DATLENH_2011.BeginCell.Split(',');
                                    for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (dataTable.Rows[i][column[0]].ToString() != "" && dataTable.Rows[i][column[0]].ToString() != "Tổng cộng")
                                        {
                                            TinhHinhDatLenh2011 th_hnx = new TinhHinhDatLenh2011();
                                            //Symbol,NumberofBids_QT,BidVolume_QT,NumberofOffers_QT
                                            //,OfferVolume_QT,Difference_QT,NumberofBids_NT
                                            //,BidVolume_NT,NumberofOffers_NT,OfferVolume_NT
                                            //,Difference_NT,SLDatMua,KLDatMua,SLDatBan,KLDatBan,Trangding_Date

                                            th_hnx.Symbol = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                th_hnx.NumberofBids_QT = 0;

                                            }
                                            else { th_hnx.NumberofBids_QT = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                th_hnx.BidVolume_QT = 0;

                                            }
                                            else { th_hnx.BidVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                th_hnx.NumberofOffers_QT = 0;

                                            }
                                            else { th_hnx.NumberofOffers_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                th_hnx.OfferVolume_QT = 0;

                                            }
                                            else
                                            {
                                                th_hnx.OfferVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                th_hnx.Difference_QT = 0;

                                            }
                                            else { th_hnx.Difference_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                th_hnx.NumberofBids_NT = 0;

                                            }
                                            else { th_hnx.NumberofBids_NT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                            //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                th_hnx.BidVolume_NT = 0;

                                            }
                                            else { th_hnx.BidVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                th_hnx.NumberofOffers_NT = 0;

                                            }
                                            else { th_hnx.NumberofOffers_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                th_hnx.OfferVolume_NT = 0;

                                            }
                                            else { th_hnx.OfferVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                th_hnx.Difference_NT = 0;

                                            }
                                            else { th_hnx.Difference_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                th_hnx.SLDatMua = 0;

                                            }
                                            else { th_hnx.SLDatMua = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                th_hnx.KLDatMua = 0;

                                            }
                                            else { th_hnx.KLDatMua = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                th_hnx.SLDatBan = 0;

                                            }
                                            else { th_hnx.SLDatBan = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                th_hnx.KLDatBan = 0;

                                            }
                                            else { th_hnx.KLDatBan = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            th_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, th_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TH_DATLENH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //NDTNN_2011
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.NDTNN_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenNDTNN_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.NDTNN_2011.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                        {
                                            NDTNN2011 nt_hnx = new NDTNN2011();
                                            //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,

                                            nt_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            nt_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                nt_hnx.KLCKMAX = 0;

                                            }
                                            else { nt_hnx.KLCKMAX = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                nt_hnx.KLMUA_QT = 0;

                                            }
                                            else { nt_hnx.KLMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                nt_hnx.GTMUA_QT = 0;

                                            }
                                            else { nt_hnx.GTMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                nt_hnx.KLBAN_QT = 0;

                                            }
                                            else
                                            {
                                                nt_hnx.KLBAN_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                nt_hnx.GIATRI_QT = 0;

                                            }
                                            else { nt_hnx.GIATRI_QT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                nt_hnx.KLMUA_NT = 0;

                                            }
                                            else { nt_hnx.KLMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                            //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                nt_hnx.GTMUA_NT = 0;

                                            }
                                            else { nt_hnx.GTMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            //,KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                nt_hnx.KLBAN_NT = 0;

                                            }
                                            else { nt_hnx.KLBAN_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                nt_hnx.GIATRI_NT = 0;

                                            }
                                            else { nt_hnx.GIATRI_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                nt_hnx.CurrentRoom = 0;

                                            }
                                            else { nt_hnx.CurrentRoom = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                nt_hnx.KLLH = 0;

                                            }
                                            else { nt_hnx.KLLH = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                nt_hnx.NamGiuMax = 0;

                                            }
                                            else { nt_hnx.NamGiuMax = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                            {
                                                nt_hnx.KLNDTN = 0;

                                            }
                                            else { nt_hnx.KLNDTN = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                            nt_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, nt_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.NDTNN_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//

                        //KQGD_2011 KQGD chi tiet
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGD_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenKQGD_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.KQGD_2011.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                        {
                                            KQGDCHITIET2011 ct_hnx = new KQGDCHITIET2011();
                                            //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighPrice,

                                            ct_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            ct_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                ct_hnx.BasicPrice = 0;

                                            }
                                            else { ct_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                ct_hnx.OpenPrice = 0;

                                            }
                                            else { ct_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                ct_hnx.ClosePrice = 0;

                                            }
                                            else { ct_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                ct_hnx.HighPrice = 0;

                                            }
                                            else
                                            {
                                                ct_hnx.HighPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            // //LowPrice,AveragePrice,NetChange,

                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                ct_hnx.LowPrice = 0;

                                            }
                                            else { ct_hnx.LowPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                ct_hnx.AveragePrice = 0;

                                            }
                                            else { ct_hnx.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                ct_hnx.NetChange = 0;

                                            }
                                            else { ct_hnx.NetChange = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                            //,Value_TC,GiaTriTT,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_BG = 0;

                                            }
                                            else { ct_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                ct_hnx.Value_BG = 0;

                                            }
                                            else { ct_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                ct_hnx.AveragePrice_TT = 0;

                                            }
                                            else { ct_hnx.AveragePrice_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_TT = 0;

                                            }
                                            else { ct_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                ct_hnx.Value_TT = 0;

                                            }
                                            else { ct_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_TC = 0;

                                            }
                                            else { ct_hnx.Volume_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                            {
                                                ct_hnx.Value_TC = 0;

                                            }
                                            else { ct_hnx.Value_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                            if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                            {
                                                ct_hnx.GiaTriTT = 0;

                                            }
                                            else { ct_hnx.GiaTriTT = Convert.ToDouble(dataTable.Rows[i][column[16]]); }


                                            ct_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, ct_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//

                        //KQGDTH_2011 KQGDTH
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGDTH_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenKQGDTH_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.KQGDTH_2011.BeginCell.Split(',');
                                    for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                    {

                                        KQGDTH2011 KQ_hnx = new KQGDTH2011();
                                        //
                                        //TypeName,Volume_BG,Value_BG,Weight_BG,Volume_TT,Value_TT,

                                        KQ_hnx.TypeName = dataTable.Rows[i][column[0]].ToString();
                                        if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_BG = 0;

                                        }
                                        else { KQ_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_BG = 0;

                                        }
                                        else { KQ_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_BG = 0;

                                        }
                                        else { KQ_hnx.Weight_BG = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                        //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_TT = 0;

                                        }
                                        else
                                        {
                                            KQ_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                        }
                                        //  //Weight_TT,Volume_MT,Value_MT,Weight_MT,Trangding_Date


                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_TT = 0;

                                        }
                                        else { KQ_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_TT = 0;

                                        }
                                        else { KQ_hnx.Weight_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_MT = 0;

                                        }
                                        else { KQ_hnx.Volume_MT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                        //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                        //,Value_TC,GiaTriTT,Trangding_Date
                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_MT = 0;

                                        }
                                        else { KQ_hnx.Value_MT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_MT = 0;

                                        }
                                        else { KQ_hnx.Weight_MT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }

                                        KQ_hnx.Trangding_Date = dateFile;
                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, KQ_hnx, null, null, null, null, null, null, null, null, null, null, null);
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGDTH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//

                        //Top_2011
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;

                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.Top_2011.SheetName];
                                    if (dataSet != null)
                                    {
                                        if (dataSet.Rows[3]["Column0"].ToString() == "")
                                        {

                                            //Top10CK_GTGDL

                                            DataTable dataTable;
                                            if (dataSet.Rows[2]["Column1"].ToString() == "")
                                            {
                                                dataTable = configTable.DatTenTop10CK_GTGDL(dataSet);
                                            }
                                            else
                                            {
                                                dataTable = configTable.DatTenTop10CK_GTGDL_2(dataSet);
                                            }


                                            string[] column = configs.HNX_File_2011.Top_2011.Top10CK_GTGDL.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i][column[0]].ToString() != "" && float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {
                                                    Top10CK_GTGDL gtgdl = new Top10CK_GTGDL();
                                                    //Symbol,ValueN,WeightN,Trangding_Date


                                                    gtgdl.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                    if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                    {
                                                        gtgdl.ValueN = 0;

                                                    }
                                                    else { gtgdl.ValueN = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {
                                                        gtgdl.WeightN = 0;

                                                    }
                                                    else { gtgdl.WeightN = Convert.ToDouble(dataTable.Rows[i][column[2]]); }


                                                    gtgdl.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_GTGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                            //Top10CK_KLGDL
                                            DataTable dataTable1 = configTable.DatTenTop10CK_KLGDL(dataSet);



                                            string[] column1 = configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable1.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable1.Rows[i][column1[0]].ToString(), out view) && float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                                {
                                                    Top10CK_KLGDL klgdl = new Top10CK_KLGDL();
                                                    //Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date


                                                    klgdl.Symbol = dataTable1.Rows[i][column1[0]].ToString();
                                                    if (!float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                                    {
                                                        klgdl.AvePrice = 0;

                                                    }
                                                    else { klgdl.AvePrice = Convert.ToDouble(dataTable1.Rows[i][column1[1]]); }
                                                    if (!float.TryParse(dataTable1.Rows[i][column1[2]].ToString(), out view))
                                                    {
                                                        klgdl.Volume = 0;

                                                    }
                                                    else { klgdl.Volume = Convert.ToDouble(dataTable1.Rows[i][column1[2]]); }

                                                    if (!float.TryParse(dataTable1.Rows[i][column1[3]].ToString(), out view))
                                                    {
                                                        klgdl.PhanTram = 0;

                                                    }
                                                    else { klgdl.PhanTram = Convert.ToDouble(dataTable1.Rows[i][column1[3]]); }
                                                    if (!float.TryParse(dataTable1.Rows[i][column1[4]].ToString(), out view))
                                                    {
                                                        klgdl.WeightN = 0;

                                                    }
                                                    else { klgdl.WeightN = Convert.ToDouble(dataTable1.Rows[i][column1[4]]); }


                                                    klgdl.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CP_GTNYL
                                            DataTable dataTable2;
                                            if (dataSet.Rows[2]["Column13"].ToString() == "")
                                            {
                                                dataTable2 = configTable.DatTenTop10CP_GTNYL(dataSet);
                                            }
                                            else
                                            {
                                                dataTable2 = configTable.DatTenTop10CP_GTNYL_2(dataSet);
                                            }


                                            string[] column2 = configs.HNX_File_2011.Top_2011.Top10CP_GTNYL.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable2.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable2.Rows[i][column2[0]].ToString(), out view) && dataTable2.Rows[i][column2[0]].ToString() != "" && float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                                {
                                                    Top10CP_GTNYL gtnyl = new Top10CP_GTNYL();
                                                    //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date


                                                    gtnyl.Symbol = dataTable2.Rows[i][column2[0]].ToString();
                                                    if (!float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                                    {
                                                        gtnyl.AvePrice = 0;

                                                    }
                                                    else { gtnyl.AvePrice = Convert.ToDouble(dataTable2.Rows[i][column2[1]]); }
                                                    if (!float.TryParse(dataTable2.Rows[i][column2[2]].ToString(), out view))
                                                    {
                                                        gtnyl.Volume = 0;

                                                    }
                                                    else { gtnyl.Volume = Convert.ToDouble(dataTable2.Rows[i][column2[2]]); }

                                                    if (!float.TryParse(dataTable2.Rows[i][column2[3]].ToString(), out view))
                                                    {
                                                        gtnyl.GiaTriNY = 0;

                                                    }
                                                    else { gtnyl.GiaTriNY = Convert.ToDouble(dataTable2.Rows[i][column2[3]]); }


                                                    gtnyl.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, gtnyl, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CP_GTNYL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CK_TANGGIA
                                            DataTable dataTable3;
                                            if (dataSet.Rows[2]["Column21"].ToString() == "")
                                            {
                                                dataTable3 = configTable.DatTenTop10CK_TANGGIA(dataSet);
                                            }
                                            else
                                            {
                                                dataTable3 = configTable.DatTenTop10CK_TANGGIA_2(dataSet);
                                            }


                                            string[] column3 = configs.HNX_File_2011.Top_2011.Top10CK_TANGGIA.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable3.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable3.Rows[i][column3[0]].ToString(), out view) && float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                                {
                                                    Top10CK_TANGGIA tg = new Top10CK_TANGGIA();
                                                    //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date


                                                    tg.Symbol = dataTable3.Rows[i][column3[0]].ToString();
                                                    if (!float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                                    {
                                                        tg.AvePrice = 0;

                                                    }
                                                    else { tg.AvePrice = Convert.ToDouble(dataTable3.Rows[i][column3[1]]); }
                                                    if (!float.TryParse(dataTable3.Rows[i][column3[2]].ToString(), out view))
                                                    {
                                                        tg.TyLeTang = 0;

                                                    }
                                                    else { tg.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column3[2]]); }

                                                    if (!float.TryParse(dataTable3.Rows[i][column3[3]].ToString(), out view))
                                                    {
                                                        tg.KLGD = 0;

                                                    }
                                                    else { tg.KLGD = Convert.ToDouble(dataTable3.Rows[i][column3[3]]); }


                                                    tg.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, tg, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_TANGGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CK_GIAMGIA
                                            DataTable dataTable4;
                                            if (dataSet.Rows[2]["Column27"].ToString() == "")
                                            {
                                                dataTable4 = configTable.DatTenTop10CK_GIAMGIA(dataSet);
                                            }
                                            else
                                            {
                                                dataTable4 = configTable.DatTenTop10CK_GIAMGIA_2(dataSet);

                                            }


                                            string[] column4 = configs.HNX_File_2011.Top_2011.Top10CK_GIAMGIA.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable4.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable4.Rows[i][column4[0]].ToString(), out view) && dataTable4.Rows[i][column4[0]].ToString() != "" && float.TryParse(dataTable4.Rows[i][column4[1]].ToString(), out view))
                                                {
                                                    Top10CK_GIAMGIA gg = new Top10CK_GIAMGIA();
                                                    //Symbol,AvePrice,MucGiam,TyLeTang,Trangding_Date


                                                    gg.Symbol = dataTable4.Rows[i][column4[0]].ToString();
                                                    if (!float.TryParse(dataTable4.Rows[i][column4[1]].ToString(), out view))
                                                    {
                                                        gg.AvePrice = 0;

                                                    }
                                                    else { gg.AvePrice = Convert.ToDouble(dataTable4.Rows[i][column4[1]]); }
                                                    if (!float.TryParse(dataTable4.Rows[i][column4[2]].ToString(), out view))
                                                    {
                                                        gg.MucGiam = 0;

                                                    }
                                                    else { gg.MucGiam = Convert.ToDouble(dataTable4.Rows[i][column4[2]]); }

                                                    if (!float.TryParse(dataTable4.Rows[i][column4[3]].ToString(), out view))
                                                    {
                                                        gg.TyLeTang = 0;

                                                    }
                                                    else { gg.TyLeTang = Convert.ToDouble(dataTable4.Rows[i][column4[3]]); }


                                                    gg.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, gg, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_GIAMGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }


                                            //Chi_Tieu_2011
                                            DataTable dataTable5;
                                            if (dataSet.Rows[2]["Column34"].ToString() == "")
                                            {
                                                dataTable5 = configTable.DatTenChi_Tieu_2011(dataSet);


                                            }
                                            else
                                            {
                                                dataTable5 = configTable.DatTenChi_Tieu_2011_2(dataSet);
                                            }


                                            string[] column5 = configs.HNX_File_2011.Top_2011.Chi_Tieu_2011.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable5.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable5.Rows[i][column5[0]].ToString(), out view) && float.TryParse(dataTable5.Rows[i][column5[1]].ToString(), out view))
                                                {
                                                    Chi_Tieu_2011 ct = new Chi_Tieu_2011();
                                                    //Chi_Tieu,CPNY,CP_DKGD_UPCOM,Trangding_Date


                                                    ct.Chi_Tieu = dataTable5.Rows[i][column5[0]].ToString();
                                                    if (!float.TryParse(dataTable5.Rows[i][column5[1]].ToString(), out view))
                                                    {
                                                        ct.CPNY = 0;

                                                    }
                                                    else { ct.CPNY = Convert.ToDouble(dataTable5.Rows[i][column5[1]]); }
                                                    if (!float.TryParse(dataTable5.Rows[i][column5[2]].ToString(), out view))
                                                    {
                                                        ct.CP_DKGD_UPCOM = 0;

                                                    }
                                                    else { ct.CP_DKGD_UPCOM = Convert.ToDouble(dataTable5.Rows[i][column5[2]]); }


                                                    ct.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, null, ct, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Chi_Tieu_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CK_NDTNN
                                            DataTable dataTable6;
                                            if (dataSet.Rows[2]["Column38"].ToString() == "")
                                            {
                                                dataTable6 = configTable.DatTenTop10CK_NDTNN(dataSet);
                                            }
                                            else
                                            {
                                                dataTable6 = configTable.DatTenTop10CK_NDTNN_2(dataSet);
                                            }


                                            string[] column6 = configs.HNX_File_2011.Top_2011.Top10CK_NDTNN.BeginCell.Split(',');
                                            int kl = dataTable6.Rows.Count - 13;
                                            for (int i = 3; i < dataTable6.Rows.Count - kl; i++)
                                            {
                                                if (!float.TryParse(dataTable6.Rows[i][column6[0]].ToString(), out view) && float.TryParse(dataTable6.Rows[i][column6[1]].ToString(), out view))
                                                {
                                                    Top10CK_NDTNN ndtnn = new Top10CK_NDTNN();
                                                    //Symbol,KLMua,GTMua,KLDPNamGiu,Trangding_Date


                                                    ndtnn.Symbol = dataTable6.Rows[i][column6[0]].ToString();
                                                    if (!float.TryParse(dataTable6.Rows[i][column6[1]].ToString(), out view))
                                                    {
                                                        ndtnn.KLMua = 0;

                                                    }
                                                    else { ndtnn.KLMua = Convert.ToDouble(dataTable6.Rows[i][column6[1]]); }
                                                    if (!float.TryParse(dataTable6.Rows[i][column6[2]].ToString(), out view))
                                                    {
                                                        ndtnn.GTMua = 0;

                                                    }
                                                    else { ndtnn.GTMua = Convert.ToDouble(dataTable6.Rows[i][column6[2]]); }

                                                    if (!float.TryParse(dataTable6.Rows[i][column6[3]].ToString(), out view))
                                                    {
                                                        ndtnn.KLDPNamGiu = 0;

                                                    }
                                                    else { ndtnn.KLDPNamGiu = Convert.ToDouble(dataTable6.Rows[i][column6[3]]); }

                                                    ndtnn.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, null, null, ndtnn, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_NDTNN.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //KLGD_TOP2011_MR
                                            DataTable dataTable7;
                                            if (dataSet.Columns.Contains("Column1"))
                                            {
                                                dataTable7 = configTable.DatTenKLGD_TOP2011_MR(dataSet);
                                            }
                                            else
                                            {
                                                dataTable7 = configTable.DatTenKLGD_TOP2011_MR_2(dataSet);
                                            }


                                            string[] column7 = configs.HNX_File_2011.Top_2011.KLGD_TOP2011_MR.BeginCell.Split(',');
                                            for (int i = 16; i < dataTable7.Rows.Count - 0; i++)
                                            {

                                                KLGD_TOP2011_MR mr1 = new KLGD_TOP2011_MR();
                                                //Symbol,AvePrice,KL,GT,TangGiam,KLGD_NgayTruoc,Trangding_Date


                                                mr1.Symbol = dataTable7.Rows[i][column7[0]].ToString();
                                                if (!float.TryParse(dataTable7.Rows[i][column7[1]].ToString(), out view))
                                                {
                                                    mr1.AvePrice = 0;

                                                }
                                                else { mr1.AvePrice = Convert.ToDouble(dataTable7.Rows[i][column7[1]]); }
                                                if (!float.TryParse(dataTable7.Rows[i][column7[2]].ToString(), out view))
                                                {
                                                    mr1.KL = 0;

                                                }
                                                else { mr1.KL = Convert.ToDouble(dataTable7.Rows[i][column7[2]]); }

                                                if (!float.TryParse(dataTable7.Rows[i][column7[3]].ToString(), out view))
                                                {
                                                    mr1.GT = 0;

                                                }
                                                else { mr1.GT = Convert.ToDouble(dataTable7.Rows[i][column7[3]]); }

                                                if (!float.TryParse(dataTable7.Rows[i][column7[4]].ToString(), out view))
                                                {
                                                    mr1.TangGiam = 0;

                                                }
                                                else { mr1.TangGiam = Convert.ToDouble(dataTable7.Rows[i][column7[4]]); }

                                                if (!float.TryParse(dataTable7.Rows[i][column7[5]].ToString(), out view))
                                                {
                                                    mr1.KLGD_NgayTruoc = 0;

                                                }
                                                else { mr1.KLGD_NgayTruoc = Convert.ToDouble(dataTable7.Rows[i][column7[5]]); }

                                                mr1.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, null, null, null, mr1, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.KLGD_TOP2011_MR.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //GTGD_TOP2011_MR
                                            DataTable dataTable8;
                                            if (dataSet.Columns.Contains("Column14"))
                                            {
                                                dataTable8 = configTable.DatTenGTGD_TOP2011_MR(dataSet);
                                            }
                                            else { dataTable8 = configTable.DatTenGTGD_TOP2011_MR_2(dataSet); }


                                            string[] column8 = configs.HNX_File_2011.Top_2011.GTGD_TOP2011_MR.BeginCell.Split(',');
                                            for (int i = 16; i < dataTable8.Rows.Count - 0; i++)
                                            {

                                                GTGD_TOP2011_MR mr2 = new GTGD_TOP2011_MR();

                                                mr2.Symbol = dataTable8.Rows[i][column8[0]].ToString();
                                                if (!float.TryParse(dataTable8.Rows[i][column8[1]].ToString(), out view))
                                                {
                                                    mr2.AvePrice = 0;

                                                }
                                                else { mr2.AvePrice = Convert.ToDouble(dataTable8.Rows[i][column8[1]]); }
                                                if (!float.TryParse(dataTable8.Rows[i][column8[2]].ToString(), out view))
                                                {
                                                    mr2.KLGD = 0;

                                                }
                                                else { mr2.KLGD = Convert.ToDouble(dataTable8.Rows[i][column8[2]]); }

                                                if (!float.TryParse(dataTable8.Rows[i][column8[3]].ToString(), out view))
                                                {
                                                    mr2.KLNY = 0;

                                                }
                                                else { mr2.KLNY = Convert.ToDouble(dataTable8.Rows[i][column8[3]]); }

                                                if (!float.TryParse(dataTable8.Rows[i][column8[4]].ToString(), out view))
                                                {
                                                    mr2.GTNY_Trieu = 0;

                                                }
                                                else { mr2.GTNY_Trieu = Convert.ToDouble(dataTable8.Rows[i][column8[4]]); }

                                                if (!float.TryParse(dataTable8.Rows[i][column8[5]].ToString(), out view))
                                                {
                                                    mr2.GTNY_Dong = 0;

                                                }
                                                else { mr2.GTNY_Dong = Convert.ToDouble(dataTable8.Rows[i][column8[5]]); }

                                                mr2.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, null, null, null, null, mr2, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.GTGD_TOP2011_MR.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //TangGiam_TOP2011_MR
                                            DataTable dataTable9;
                                            if (dataSet.Columns.Contains("Column22"))
                                            {

                                                if (!float.TryParse(dataSet.Rows[16]["Column19"].ToString(), out view) && dataSet.Rows[16]["Column19"].ToString() == "")
                                                {
                                                    dataTable9 = configTable.DatTenTangGiam_TOP2011_MR(dataSet);
                                                }
                                                else
                                                {
                                                    dataTable9 = configTable.DatTenTangGiam_TOP2011_MR_SS(dataSet);
                                                }
                                            }
                                            else
                                            {
                                                dataTable9 = configTable.DatTenTangGiam_TOP2011_MR_2(dataSet);
                                            }


                                            string[] column9 = configs.HNX_File_2011.Top_2011.TangGiam_TOP2011_MR.BeginCell.Split(',');
                                            for (int i = 16; i < dataTable9.Rows.Count - 0; i++)
                                            {

                                                TangGiam_TOP2011_MR mr3 = new TangGiam_TOP2011_MR();
                                                //Symbol,AvePrice,MucTang,PTTangGiam,KLGD,CEILINGPRICE,ChenhLechTran,
                                                //FLOORPRICES,ChenhLechSan,Trangding_Date
                                                mr3.Symbol = dataTable9.Rows[i][column9[0]].ToString();
                                                if (!float.TryParse(dataTable9.Rows[i][column9[1]].ToString(), out view))
                                                {
                                                    mr3.AvePrice = 0;

                                                }
                                                else { mr3.AvePrice = Convert.ToDouble(dataTable9.Rows[i][column9[1]]); }
                                                if (!float.TryParse(dataTable9.Rows[i][column9[2]].ToString(), out view))
                                                {
                                                    mr3.MucTang = 0;

                                                }
                                                else { mr3.MucTang = Convert.ToDouble(dataTable9.Rows[i][column9[2]]); }

                                                if (!float.TryParse(dataTable9.Rows[i][column9[3]].ToString(), out view))
                                                {
                                                    mr3.PTTangGiam = 0;

                                                }
                                                else { mr3.PTTangGiam = Convert.ToDouble(dataTable9.Rows[i][column9[3]]); }

                                                if (!float.TryParse(dataTable9.Rows[i][column9[4]].ToString(), out view))
                                                {
                                                    mr3.KLGD = 0;

                                                }
                                                else { mr3.KLGD = Convert.ToDouble(dataTable9.Rows[i][column9[4]]); }

                                                if (!float.TryParse(dataTable9.Rows[i][column9[5]].ToString(), out view))
                                                {
                                                    mr3.CEILINGPRICE = 0;

                                                }
                                                else { mr3.CEILINGPRICE = Convert.ToDouble(dataTable9.Rows[i][column9[5]]); }

                                                if (!float.TryParse(dataTable9.Rows[i][column9[6]].ToString(), out view))
                                                {
                                                    mr3.ChenhLechTran = 0;

                                                }
                                                else { mr3.ChenhLechTran = Convert.ToDouble(dataTable9.Rows[i][column9[6]]); }

                                                if (!float.TryParse(dataTable9.Rows[i][column9[7]].ToString(), out view))
                                                {
                                                    mr3.FLOORPRICES = 0;

                                                }
                                                else { mr3.FLOORPRICES = Convert.ToDouble(dataTable9.Rows[i][column9[7]]); }
                                                if (!float.TryParse(dataTable9.Rows[i][column9[8]].ToString(), out view))
                                                {
                                                    mr3.ChenhLechSan = 0;

                                                }
                                                else { mr3.ChenhLechSan = Convert.ToDouble(dataTable9.Rows[i][column9[8]]); }

                                                mr3.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, null, null, null, null, null, mr3, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.TangGiam_TOP2011_MR.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //CKNTDNN_TOP2011_MR
                                            bool columnExists = dataSet.Columns.Contains("Column43");
                                            DataTable dataTable10;
                                            if (columnExists)
                                            {
                                                dataTable10 = configTable.DatTenCKNTDNN_TOP2011_MR(dataSet);
                                            }
                                            else
                                            {
                                                dataTable10 = configTable.DatTenCKNTDNN_TOP2011_MR_2(dataSet);
                                            }

                                            string[] column10 = configs.HNX_File_2011.Top_2011.CKNTDNN_TOP2011_MR.BeginCell.Split(',');
                                            for (int i = 16; i < dataTable10.Rows.Count - 0; i++)
                                            {

                                                CKNTDNN_TOP2011_MR mr4 = new CKNTDNN_TOP2011_MR();
                                                //Symbol,KLMua,GTMua,KLDPNamGiu,Trangding_Date
                                                mr4.Symbol = dataTable10.Rows[i][column10[0]].ToString();
                                                if (!float.TryParse(dataTable10.Rows[i][column10[1]].ToString(), out view))
                                                {
                                                    mr4.KLMua = 0;

                                                }
                                                else { mr4.KLMua = Convert.ToDouble(dataTable10.Rows[i][column10[1]]); }
                                                if (!float.TryParse(dataTable10.Rows[i][column10[2]].ToString(), out view))
                                                {
                                                    mr4.GTMua = 0;

                                                }
                                                else { mr4.GTMua = Convert.ToDouble(dataTable10.Rows[i][column10[2]]); }

                                                if (columnExists)
                                                {
                                                    if (!float.TryParse(dataTable10.Rows[i][column10[3]].ToString(), out view))
                                                    {
                                                        mr4.KLBan = 0;

                                                    }
                                                    else { mr4.KLBan = Convert.ToDouble(dataTable10.Rows[i][column10[3]]); }
                                                    if (!float.TryParse(dataTable10.Rows[i][column10[4]].ToString(), out view))
                                                    {
                                                        mr4.GTBan = 0;

                                                    }
                                                    else { mr4.GTBan = Convert.ToDouble(dataTable10.Rows[i][column10[4]]); }



                                                    if (!float.TryParse(dataTable10.Rows[i][column10[5]].ToString(), out view))
                                                    {
                                                        mr4.KLDPNamGiu = 0;

                                                    }
                                                    else { mr4.KLDPNamGiu = Convert.ToDouble(dataTable10.Rows[i][column10[5]]); }


                                                }
                                                else
                                                {
                                                    mr4.KLBan = 0;
                                                    mr4.GTBan = 0;

                                                    if (!float.TryParse(dataTable10.Rows[i][column10[3]].ToString(), out view))
                                                    {
                                                        mr4.KLDPNamGiu = 0;
                                                    }
                                                    else { mr4.KLDPNamGiu = Convert.ToDouble(dataTable10.Rows[i][column10[3]]); }

                                                }

                                                mr4.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, mr4);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);


                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.CKNTDNN_TOP2011_MR.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                            //Console.WriteLine("File: " + filePath);


                                        }
                                        else
                                        {
                                            //Top10CK_GTGDL

                                            DataTable dataTable = configTable.DatTenTop10CK_GTGDL_2010(dataSet);


                                            string[] column = configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Column3"].ToString() == "" && dataTable.Rows[i]["Column4"].ToString() != "")
                                                {
                                                    Top10CK_GTGDL gtgdl = new Top10CK_GTGDL();
                                                    //Symbol,ValueN,WeightN,Trangding_Date


                                                    gtgdl.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                    if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                    {
                                                        gtgdl.ValueN = 0;

                                                    }
                                                    else { gtgdl.ValueN = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {
                                                        gtgdl.WeightN = 0;

                                                    }
                                                    else { gtgdl.WeightN = Convert.ToDouble(dataTable.Rows[i][column[2]]); }


                                                    gtgdl.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                            //Top10CK_KLGDL
                                            DataTable dataTable1 = configTable.DatTenTop10CK_KLGDL_2010(dataSet);



                                            string[] column1 = configs.HNX_File_2010.Top_2010.Top10CK_KLGDL.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable1.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable1.Rows[i][column1[0]].ToString(), out view) && dataTable1.Rows[i][column1[1]].ToString() != "" && dataTable1.Rows[i][column1[0]].ToString() != "")
                                                {
                                                    Top10CK_KLGDL klgdl = new Top10CK_KLGDL();
                                                    //Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date


                                                    klgdl.Symbol = dataTable1.Rows[i][column1[0]].ToString();

                                                    klgdl.AvePrice = 0;


                                                    if (!float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                                    {
                                                        klgdl.Volume = 0;

                                                    }
                                                    else { klgdl.Volume = Convert.ToDouble(dataTable1.Rows[i][column1[1]]); }

                                                    klgdl.PhanTram = 0;


                                                    if (!float.TryParse(dataTable1.Rows[i][column1[2]].ToString(), out view))
                                                    {
                                                        klgdl.WeightN = 0;

                                                    }
                                                    else { klgdl.WeightN = Convert.ToDouble(dataTable1.Rows[i][column1[2]]); }


                                                    klgdl.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CP_GTNYL
                                            DataTable dataTable2 = configTable.DatTenTop10CP_GTNYL_2010_1(dataSet);



                                            string[] column2 = configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable2.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable2.Rows[i][column2[0]].ToString(), out view) && dataTable2.Rows[i][column2[0]].ToString() != "" && float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                                {
                                                    Top10CP_CLGMAX gtnyl = new Top10CP_CLGMAX();
                                                    //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date


                                                    gtnyl.Symbol = dataTable2.Rows[i][column2[0]].ToString();
                                                    if (!float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                                    {
                                                        gtnyl.HighPrice = 0;

                                                    }
                                                    else { gtnyl.HighPrice = Convert.ToDouble(dataTable2.Rows[i][column2[1]]); }
                                                    if (!float.TryParse(dataTable2.Rows[i][column2[2]].ToString(), out view))
                                                    {
                                                        gtnyl.LowPrice = 0;

                                                    }
                                                    else { gtnyl.LowPrice = Convert.ToDouble(dataTable2.Rows[i][column2[2]]); }

                                                    if (!float.TryParse(dataTable2.Rows[i][column2[3]].ToString(), out view))
                                                    {
                                                        gtnyl.TyLeChenhLech = 0;

                                                    }
                                                    else { gtnyl.TyLeChenhLech = Convert.ToDouble(dataTable2.Rows[i][column2[3]]); }


                                                    gtnyl.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, gtnyl, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CK_TANGGIA
                                            DataTable dataTable3 = configTable.DatTenTop10CK_TANGGIA_2010(dataSet);

                                            string[] column3 = configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable3.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable3.Rows[i][column3[0]].ToString(), out view) && float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                                {
                                                    Top10CK_TANGGIA2010 tg = new Top10CK_TANGGIA2010();
                                                    //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date


                                                    tg.Symbol = dataTable3.Rows[i][column3[0]].ToString();
                                                    if (!float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                                    {
                                                        tg.AvePrice = 0;

                                                    }
                                                    else { tg.AvePrice = Convert.ToDouble(dataTable3.Rows[i][column3[1]]); }
                                                    if (!float.TryParse(dataTable3.Rows[i][column3[2]].ToString(), out view))
                                                    {
                                                        tg.MucTang = 0;

                                                    }
                                                    else { tg.MucTang = Convert.ToDouble(dataTable3.Rows[i][column3[2]]); }

                                                    if (!float.TryParse(dataTable3.Rows[i][column3[3]].ToString(), out view))
                                                    {
                                                        tg.TyLeTang = 0;

                                                    }
                                                    else { tg.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column3[3]]); }


                                                    tg.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2010(tg, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }

                                            //Top10CK_GIAMGIA
                                            DataTable dataTable4 = configTable.DatTenTop10CK_GIAMGIA_2010(dataSet);



                                            string[] column4 = configs.HNX_File_2010.Top_2010.Top10CK_GIAMGIA.BeginCell.Split(',');
                                            for (int i = 3; i < dataTable4.Rows.Count - 0; i++)
                                            {
                                                if (!float.TryParse(dataTable4.Rows[i][column4[0]].ToString(), out view) && dataTable4.Rows[i][column4[0]].ToString() != "" && float.TryParse(dataTable4.Rows[i][column4[1]].ToString(), out view))
                                                {
                                                    Top10CK_GIAMGIA gg = new Top10CK_GIAMGIA();
                                                    //Symbol,AvePrice,MucGiam,TyLeTang,Trangding_Date


                                                    gg.Symbol = dataTable4.Rows[i][column4[0]].ToString();
                                                    if (!float.TryParse(dataTable4.Rows[i][column4[1]].ToString(), out view))
                                                    {
                                                        gg.AvePrice = 0;

                                                    }
                                                    else { gg.AvePrice = Convert.ToDouble(dataTable4.Rows[i][column4[1]]); }
                                                    if (!float.TryParse(dataTable4.Rows[i][column4[2]].ToString(), out view))
                                                    {
                                                        gg.MucGiam = 0;

                                                    }
                                                    else { gg.MucGiam = Convert.ToDouble(dataTable4.Rows[i][column4[2]]); }

                                                    if (!float.TryParse(dataTable4.Rows[i][column4[3]].ToString(), out view))
                                                    {
                                                        gg.TyLeTang = 0;

                                                    }
                                                    else { gg.TyLeTang = Convert.ToDouble(dataTable4.Rows[i][column4[3]]); }


                                                    gg.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, gg, null, null, null, null, null, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }

                                            }

                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GIAMGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                            //Console.WriteLine("File: " + filePath);
                                        }


                                    }



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                    }



                }

                if (match_4.Success || match_4_1.Success)
                {
                    DateTime dateFile;
                    string thitruong;
                    if (match_4.Success)
                    {
                        string dateString = match_4.Groups[1].Value;
                        thitruong = match_4.Groups[2].Value;
                        DateTime outputDate = DateTime.ParseExact(dateString, "yyyy.MM.dd", CultureInfo.InvariantCulture);
                        string formattedDate = outputDate.ToString("yyyy-MM-dd");
                        dateFile = DateTime.ParseExact(formattedDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        string dateString = match_4_1.Groups[1].Value;
                        thitruong = match_4_1.Groups[2].Value;
                        DateTime outputDate = DateTime.ParseExact(dateString, "yyyy.MM.dd", CultureInfo.InvariantCulture);
                        string formattedDate = outputDate.ToString("yyyy-MM-dd");
                        dateFile = DateTime.ParseExact(formattedDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    }

                    switch (thitruong)
                    {
                        case ConfigApp.NY_4:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenUPCoM_GDNDTNN_Phien(dataSet);
                                        float view;
                                        DataTable dataTable;

                                        dataTable = dataSetX.Tables[configs.UPCoM_GDNDTNN_Phien.SheetName];
                                        if (dataTable == null)
                                        {
                                            dataTable = dataSetX.Tables["Sheet1"];
                                        }
                                        string[] column = configs.UPCoM_GDNDTNN_Phien.BeginCell.Split(',');
                                        for (int i = 5; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                UPCoM_GDNDTNN_Phien cc = new UPCoM_GDNDTNN_Phien();

                                                //STT,Symbol,KLMUA_KL,GTMUA_KL,KLBAN_KL

                                                cc.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc.KLMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc.KLMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc.GTMUA_KL = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc.GTMUA_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc.KLBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc.KLBAN_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc.GTBAN_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc.GTBAN_KL = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc.KLMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc.KLMUA_TT = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc.GTMUA_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc.GTMUA_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc.KLBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc.KLBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc.GTBAN_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc.GTBAN_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc.KLMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc.KLMUA_TC = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc.GTMUA_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc.GTMUA_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {

                                                    cc.KLBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                }
                                                else { cc.KLBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {

                                                    cc.GTBAN_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                }
                                                else { cc.GTBAN_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {

                                                    cc.KLCK_MAX = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                }
                                                else { cc.KLCK_MAX = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {

                                                    cc.KLCK_NDTNN = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                }
                                                else { cc.KLCK_NDTNN = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {

                                                    cc.KLCK_CDPNG = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                }
                                                else { cc.KLCK_CDPNG = 0; }

                                                cc.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM_2013(cc, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_GDNDTNN_Phien.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        case ConfigApp.NY_2:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX_1 = configTable.DatTenUPCOMEDO2_1(dataSet);
                                        //MAI LÀM TIẾP
                                        var dataSetX_2 = configTable.DatTenUPCOMEOD02_2(dataSet);

                                        float view;
                                        DataTable dataTable = dataSetX_1.Tables[configs.UPCoM_TT_Phien.SheetName];
                                        DataTable dataTable2 = dataSetX_2.Tables[configs.UPCoM_TT_Phien.SheetName];

                                        string[] column = configs.UPCoM_TT_Phien.Data_Table_Chi_Tieu.BeginCell.Split(',');
                                        string[] column_2 = configs.UPCoM_TT_Phien.Data_Table_Top10_CPGDT.BeginCell.Split(',');
                                        string[] column_3 = configs.UPCoM_TT_Phien.Data_Table_Top10_CPTPRICE.BeginCell.Split(',');
                                        string[] column_4 = configs.UPCoM_TT_Phien.Data_Table_Top10_KLGDM.BeginCell.Split(',');
                                        string[] column_5 = configs.UPCoM_TT_Phien.Data_Table_Top10_CPGIAMGIA.BeginCell.Split(',');
                                        for (int i = 4; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (dataTable.Rows[i][column[0]].ToString() != "")
                                            {

                                                Chi_Tieu_UPCOM ct = new Chi_Tieu_UPCOM();

                                                //Chi_Tieu,Don_Vi,So_Lieu,Trangding_Date

                                                ct.Chi_Tieu = dataTable.Rows[i][column[0]].ToString();
                                                ct.Don_Vi = "";
                                                if (float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {

                                                    ct.So_Lieu = Convert.ToDouble(dataTable.Rows[i][column[1]]);

                                                }
                                                else { ct.So_Lieu = 0; }

                                                ct.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, ct, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_TT_Phien.Data_Table_Chi_Tieu.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        //==============================================================//
                                        for (int i = 5; i < dataTable2.Rows.Count - 18; i++)
                                        {
                                            if (dataTable2.Rows[i][column_2[0]].ToString() != "")
                                            {
                                                Top10_CPGDT cpgdt = new Top10_CPGDT();

                                                //Symbol,GTGD,TyTrong,Trangding_Date

                                                cpgdt.Symbol = dataTable2.Rows[i][column_2[0]].ToString();
                                                if (float.TryParse(dataTable2.Rows[i][column_2[1]].ToString(), out view))
                                                {
                                                    cpgdt.GTGD = Convert.ToDouble(dataTable2.Rows[i][column_2[1]]);
                                                }
                                                else { cpgdt.GTGD = 0; }
                                                if (float.TryParse(dataTable2.Rows[i][column_2[2]].ToString(), out view))
                                                {

                                                    cpgdt.TyTrong = Convert.ToDouble(dataTable2.Rows[i][column_2[2]]);

                                                }
                                                else { cpgdt.TyTrong = 0; }

                                                cpgdt.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, cpgdt, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_TT_Phien.Data_Table_Top10_CPGDT.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        //=============================================================================//
                                        var dataSetX_3 = configTable.DatTenUPCOMEDO2_3(dataSet);
                                        DataTable dataTable3 = dataSetX_3.Tables[configs.UPCoM_TT_Phien.SheetName];

                                        for (int i = 23; i < dataTable3.Rows.Count - 0; i++)
                                        {
                                            Top10_CPTPRICE cptprice = new Top10_CPTPRICE();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cptprice.Symbol = dataTable3.Rows[i][column_3[0]].ToString();
                                            if (float.TryParse(dataTable3.Rows[i][column_3[1]].ToString(), out view))
                                            {
                                                cptprice.MucTang = Convert.ToDouble(dataTable3.Rows[i][column_3[1]]);
                                            }
                                            else { cptprice.MucTang = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[2]].ToString(), out view))
                                            {

                                                cptprice.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column_3[2]]);

                                            }
                                            else { cptprice.TyLeTang = 0; }
                                            if (float.TryParse(dataTable3.Rows[i][column_3[3]].ToString(), out view))
                                            {

                                                cptprice.KLGD = Convert.ToDouble(dataTable3.Rows[i][column_3[3]]);

                                            }
                                            else { cptprice.KLGD = 0; }

                                            cptprice.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, cptprice, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_TT_Phien.Data_Table_Top10_CPTPRICE.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        //=======================================================================//
                                        var dataSetX_4 = configTable.DatTenUPCOMEDO2_4(dataSet);
                                        DataTable dataTable4 = dataSetX_4.Tables[configs.UPCoM_TT_Phien.SheetName];
                                        for (int i = 5; i < dataTable4.Rows.Count - 18; i++)
                                        {
                                            Top10_KLGDM klgdm = new Top10_KLGDM();

                                            //Symbol,GTGD,TyTrong,Trangding_Date

                                            klgdm.Symbol = dataTable4.Rows[i][column_4[0]].ToString();
                                            if (float.TryParse(dataTable2.Rows[i][column_4[1]].ToString(), out view))
                                            {
                                                klgdm.KLGD = Convert.ToDouble(dataTable4.Rows[i][column_4[1]]);
                                            }
                                            else { klgdm.KLGD = 0; }
                                            if (float.TryParse(dataTable4.Rows[i][column_4[2]].ToString(), out view))
                                            {

                                                klgdm.TyTrong = Convert.ToDouble(dataTable4.Rows[i][column_4[2]]);

                                            }
                                            else { klgdm.TyTrong = 0; }

                                            klgdm.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, null, klgdm, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_TT_Phien.Data_Table_Top10_KLGDM.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }
                                        //=====================================================//
                                        var dataSetX_5 = configTable.DatTenUPCOMEDO2_5(dataSet);
                                        DataTable dataTable5 = dataSetX_5.Tables[configs.UPCoM_TT_Phien.SheetName];

                                        for (int i = 23; i < dataTable5.Rows.Count - 0; i++)
                                        {
                                            Top10_CPGIAMGIA cpgiamgia = new Top10_CPGIAMGIA();

                                            //Symbol,MucTang,TyLeTang,KLGD,Trangding_Date

                                            cpgiamgia.Symbol = dataTable5.Rows[i][column_5[0]].ToString();
                                            if (float.TryParse(dataTable5.Rows[i][column_5[1]].ToString(), out view))
                                            {
                                                cpgiamgia.MucGIAM = Convert.ToDouble(dataTable5.Rows[i][column_5[1]]);
                                            }
                                            else { cpgiamgia.MucGIAM = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[2]].ToString(), out view))
                                            {

                                                cpgiamgia.TyLeGiam = Convert.ToDouble(dataTable5.Rows[i][column_5[2]]);

                                            }
                                            else { cpgiamgia.TyLeGiam = 0; }
                                            if (float.TryParse(dataTable5.Rows[i][column_5[3]].ToString(), out view))
                                            {

                                                cpgiamgia.KLGD = Convert.ToDouble(dataTable5.Rows[i][column_5[3]]);

                                            }
                                            else { cpgiamgia.KLGD = 0; }

                                            cpgiamgia.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBUPCOM(null, null, null, null, null, null, null, null, cpgiamgia, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_TT_Phien.Data_Table_Top10_CPGIAMGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();
                                        }

                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            break;
                        case ConfigApp.NY_3:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenUPCoM_TKCC(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.UPCoM_TKCC.SheetName];

                                        string[] column = configs.UPCoM_TKCC.BeginCell.Split(',');
                                        for (int i = 4; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                UPCoM_TKCC cc = new UPCoM_TKCC();

                                                cc.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc.SLDATMUA = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc.SLDATMUA = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc.KLDATMUA = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc.KLDATMUA = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc.SLDATBAN = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc.SLDATBAN = 0; }
                                                //OpenPrice,ClosePrice,AveragePrice,TDDiem,TDPhanTram

                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc.KLDATBAN = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc.KLDATBAN = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc.CLMUABAN = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc.CLMUABAN = 0; }


                                                cc.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM_2013(null, null, cc, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_TKCC.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        case ConfigApp.NY_1:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenUPCoM_KQGD_Phien(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.UPCoM_KQGD_Phien.SheetName];

                                        string[] column = configs.UPCoM_KQGD_Phien.BeginCell.Split(',');
                                        for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                UPCoM_KQGD_Phien cc = new UPCoM_KQGD_Phien();

                                                //STT,Symbol,BasicPrice,HighestPrice,LowestPrice,


                                                cc.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc.BasicPrice = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc.HighestPrice = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc.LowestPrice = 0; }
                                                //OpenPrice,ClosePrice,AveragePrice,TDDiem,TDPhanTram

                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc.OpenPrice = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc.ClosePrice = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc.AveragePrice = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc.TDDiem = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc.TDDiem = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc.TDPhanTram = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc.TDPhanTram = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc.KLGD_KL = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc.KLGD_KL = 0; }
                                                //,KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT,KLGD_TC,TITRONG1

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc.GTGD_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc.GTGD_KL = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {

                                                    cc.KLGD_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                                }
                                                else { cc.KLGD_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {

                                                    cc.GTGD_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                                }
                                                else { cc.GTGD_TT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {

                                                    cc.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                                }
                                                else { cc.KLGD_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {

                                                    cc.TITRONG1 = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                                }
                                                else { cc.TITRONG1 = 0; }
                                                //,GTGD_TC,TITRONG2,KLCPLH,GTVHTT_GT,GTVHTT_TT,Trangding_Date
                                                if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {

                                                    cc.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                                }
                                                else { cc.GTGD_TC = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                                {

                                                    cc.TITRONG2 = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                                }
                                                else { cc.TITRONG2 = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                                {

                                                    cc.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                                }
                                                else { cc.KLCPLH = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                                {

                                                    cc.GTVHTT_GT = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                                }
                                                else { cc.GTVHTT_GT = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                                {

                                                    cc.GTVHTT_TT = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                                }
                                                else { cc.GTVHTT_TT = 0; }

                                                cc.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM_2013(null, cc, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_KQGD_Phien.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        case ConfigApp.NY_5:
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        var dataSet = reader.AsDataSet();
                                        var dataSetX = configTable.DatTenUPCoM_CPDKGD_Phien(dataSet);
                                        float view;
                                        DataTable dataTable = dataSetX.Tables[configs.UPCoM_CPDKGD_Phien.SheetName];

                                        string[] column = configs.UPCoM_CPDKGD_Phien.BeginCell.Split(',');
                                        for (int i = 4; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                UPCoM_CPDKGD_Phien cc = new UPCoM_CPDKGD_Phien();

                                                //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2010,PE,EPS2010,ROE,ROA


                                                cc.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                cc.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {

                                                    cc.KLCP_NY = Convert.ToDouble(dataTable.Rows[i][column[2]]);

                                                }
                                                else { cc.KLCP_NY = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {

                                                    cc.KLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                                }
                                                else { cc.KLCP_LH = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {

                                                    cc.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                                }
                                                else { cc.Co_Tuc_2010 = 0; }
                                                //OpenPrice,ClosePrice,AveragePrice,TDDiem,TDPhanTram

                                                if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {

                                                    cc.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                                }
                                                else { cc.PE = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {

                                                    cc.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                                }
                                                else { cc.EPS2010 = 0; }


                                                if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {

                                                    cc.ROE = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                                }
                                                else { cc.ROE = 0; }
                                                if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {

                                                    cc.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                                }
                                                else { cc.ROA = 0; }
                                                //,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date

                                                if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {

                                                    cc.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                                }
                                                else { cc.BasicPrice_KT = 0; }

                                                if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {

                                                    cc.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                                }
                                                else { cc.CeilingPrice_KT = 0; }
                                                //,KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT,KLGD_TC,TITRONG1

                                                if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {

                                                    cc.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                                }
                                                else { cc.FloorPrice_KT = 0; }

                                                cc.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBUPCOM_2013(null, null, null, cc);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_CPDKGD_Phien.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            break;
                        default:
                            break;

                    }
                }
                if (match_6.Success)
                {
                    try
                    {

                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                EBulkScript eBulkScript = new EBulkScript();
                                var dataSet = reader.AsDataSet();
                                var dataSetX = configTable.DatTenUPCoM_GDNDTNN_Phien_2011(dataSet);
                                float view;
                                DataTable dataTable = dataSetX.Tables[configs.UPCoM_GDNDTNN_2011.SheetName];

                                string[] column = configs.UPCoM_GDNDTNN_2011.BeginCell.Split(',');
                                for (int i = 4; i < dataTable.Rows.Count - 0; i++)
                                {
                                    if (dataTable.Rows[i][column[0]].ToString() != "")
                                    {
                                        UPCoM_KQGD_Phien_2011 cc = new UPCoM_KQGD_Phien_2011();

                                        //"Trangding_Date,GiaoDich,Symbol,
                                        //BasicPrice,CellingPrice,FloorPrice,


                                        cc.GiaoDich = dataTable.Rows[i][column[1]].ToString();
                                        cc.Symbol = dataTable.Rows[i][column[2]].ToString();
                                        if (float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                        {

                                            cc.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]);

                                        }
                                        else { cc.BasicPrice = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                        {

                                            cc.CellingPrice = Convert.ToDouble(dataTable.Rows[i][column[4]]);

                                        }
                                        else { cc.CellingPrice = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                        {

                                            cc.FloorPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);

                                        }
                                        else { cc.FloorPrice = 0; }
                                        //HighestPrice,LowestPrice,OpenPrice,ClosePrice,


                                        if (float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                        {

                                            cc.HighestPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]);

                                        }
                                        else { cc.HighestPrice = 0; }

                                        if (float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                        {

                                            cc.LowestPrice = Convert.ToDouble(dataTable.Rows[i][column[7]]);

                                        }
                                        else { cc.LowestPrice = 0; }


                                        if (float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                        {

                                            cc.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[8]]);

                                        }
                                        else { cc.OpenPrice = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                        {

                                            cc.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[9]]);

                                        }
                                        else { cc.ClosePrice = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                        {

                                            cc.Gia_BQ = Convert.ToDouble(dataTable.Rows[i][column[10]]);

                                        }
                                        else { cc.Gia_BQ = 0; }
                                        //Gia_BQ,KLGD_KL,GTGD_KL,HighestPrice_TT,LowestPrice_TT

                                        if (float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                        {

                                            cc.KLGD_KL = Convert.ToDouble(dataTable.Rows[i][column[11]]);

                                        }
                                        else { cc.KLGD_KL = 0; }

                                        if (float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                        {

                                            cc.GTGD_KL = Convert.ToDouble(dataTable.Rows[i][column[12]]);

                                        }
                                        else { cc.GTGD_KL = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                        {

                                            cc.HighestPrice_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]);

                                        }
                                        else { cc.HighestPrice_TT = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                        {

                                            cc.LowestPrice_TT = Convert.ToDouble(dataTable.Rows[i][column[14]]);

                                        }
                                        else { cc.LowestPrice_TT = 0; }
                                        //,KLGD_TT,GTGD_TT,KLGD_TC,GTGD_TC

                                        if (float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                        {

                                            cc.KLGD_TT = Convert.ToDouble(dataTable.Rows[i][column[15]]);

                                        }
                                        else { cc.KLGD_TT = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                        {

                                            cc.GTGD_TT = Convert.ToDouble(dataTable.Rows[i][column[16]]);

                                        }
                                        else { cc.GTGD_TT = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[17]].ToString(), out view))
                                        {

                                            cc.KLGD_TC = Convert.ToDouble(dataTable.Rows[i][column[17]]);

                                        }
                                        else { cc.KLGD_TC = 0; }

                                        if (float.TryParse(dataTable.Rows[i][column[18]].ToString(), out view))
                                        {

                                            cc.GTGD_TC = Convert.ToDouble(dataTable.Rows[i][column[18]]);

                                        }
                                        else { cc.GTGD_TC = 0; }
                                        //,Muc_VHTT,KL_DKGD,KLCPLH,KLMUA,GTMUA


                                        if (float.TryParse(dataTable.Rows[i][column[19]].ToString(), out view))
                                        {

                                            cc.Muc_VHTT = Convert.ToDouble(dataTable.Rows[i][column[19]]);

                                        }
                                        else { cc.Muc_VHTT = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[20]].ToString(), out view))
                                        {

                                            cc.KL_DKGD = Convert.ToDouble(dataTable.Rows[i][column[20]]);

                                        }
                                        else { cc.KL_DKGD = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[21]].ToString(), out view))
                                        {

                                            cc.KLCPLH = Convert.ToDouble(dataTable.Rows[i][column[21]]);

                                        }
                                        else { cc.KLCPLH = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[22]].ToString(), out view))
                                        {

                                            cc.KLMUA = Convert.ToDouble(dataTable.Rows[i][column[22]]);

                                        }
                                        else { cc.KLMUA = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[23]].ToString(), out view))
                                        {

                                            cc.GTMUA = Convert.ToDouble(dataTable.Rows[i][column[23]]);

                                        }
                                        else { cc.GTMUA = 0; }
                                        //,KLBAN,GTBAN,TongKLDPNG,KLCDPNG"
                                        if (float.TryParse(dataTable.Rows[i][column[24]].ToString(), out view))
                                        {

                                            cc.KLBAN = Convert.ToDouble(dataTable.Rows[i][column[24]]);

                                        }
                                        else { cc.KLBAN = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[25]].ToString(), out view))
                                        {

                                            cc.GTBAN = Convert.ToDouble(dataTable.Rows[i][column[25]]);

                                        }
                                        else { cc.GTBAN = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[26]].ToString(), out view))
                                        {

                                            cc.TongKLDPNG = Convert.ToDouble(dataTable.Rows[i][column[26]]);

                                        }
                                        else { cc.TongKLDPNG = 0; }
                                        if (float.TryParse(dataTable.Rows[i][column[27]].ToString(), out view))
                                        {

                                            cc.KLCDPNG = Convert.ToDouble(dataTable.Rows[i][column[27]]);

                                        }
                                        else { cc.KLCDPNG = 0; }
                                        string inputDate = dataTable.Rows[i][column[0]].ToString();
                                        DateTime date = DateTime.ParseExact(inputDate, "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
                                        string outputDate = date.ToString("yyyy-MM-dd");
                                        //  DateTime inputDate = DateTime.ParseExact(inputDateString, "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
                                        //  string outputDateString = inputDate.ToString("yyyy-MM-dd");
                                        DateTime dateFile = DateTime.ParseExact(outputDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                        cc.Trangding_Date = dateFile;
                                        eBulkScript = this.configTable.GetScriptTTCBUPCOM_2011(cc);
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);
                                    }
                                }
                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                {
                                    // exec script mssql+oracle
                                    string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.UPCoM_GDNDTNN_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                    configTable.ExecBulkScript(test);
                                    mssqlBuilder_HNX.Clear();

                                }

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("File erorr: " + filePath);
                    }

                }
                if (match_7.Success)
                {
                    string thoigian = match_7.Value;
                    DateTime dateTime = DateTime.ParseExact(thoigian, "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    string formattedDate = dateTime.ToString("yyyy-MM-dd");
                    DateTime dateFile = DateTime.ParseExact(formattedDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    DateTime dateNew = DateTime.ParseExact("2008-07-31", "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    string namefile = lastPart;
                    string miu = "KQGD tong hop";
                    string miux2 = "tong hop";
                    if (namefile.Contains(miu) || thoigian == "03.03.2008")
                    {
                        // DateTime dateTo = DateTime.ParseExact(configs.ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        try
                        {


                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];


                                    DataTable dataTable = configTable.DatTenTT_DKGD_2011_21(dataSet);


                                    //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                    string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                        {

                                            KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                            dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                dkgd_hnx.SLCP_DKGD = 0;

                                            }
                                            else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                dkgd_hnx.SLCP_LH = 0;

                                            }
                                            else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                dkgd_hnx.Co_Tuc_2010 = 0;

                                            }
                                            else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                dkgd_hnx.PE = 0;

                                            }
                                            else
                                            {
                                                dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                dkgd_hnx.EPS2010 = 0;

                                            }
                                            else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }

                                            dkgd_hnx.KLGD_10PHIEN = 0;

                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                dkgd_hnx.ROE = 0;

                                            }
                                            else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                dkgd_hnx.ROA = 0;

                                            }
                                            else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                dkgd_hnx.BasicPrice_KT = 0;

                                            }
                                            else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                dkgd_hnx.CeilingPrice_KT = 0;

                                            }
                                            else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                dkgd_hnx.FloorPrice_KT = 0;

                                            }
                                            else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                dkgd_hnx.BinhQuan = 0;

                                            }
                                            else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }

                                            dkgd_hnx.Tong = 0;
                                            dkgd_hnx.Co_Tuc_2009 = 0;

                                            dkgd_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }
                        //Tinh hinh dat lenh
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    // var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];
                                    DataTable dataSet;

                                    dataSet = reader.AsDataSet().Tables["Tinh hinh dt lenh"];

                                    DataTable dataTable = configTable.DatTenTH_DATLENH_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.TH_DATLENH_2011.BeginCell.Split(',');
                                    for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        string vi = dataTable.Rows[i][column[0]].ToString();
                                        if (dataTable.Rows[i][column[0]].ToString() != "" && vi.Length < 7)
                                        {
                                            TinhHinhDatLenh2011 th_hnx = new TinhHinhDatLenh2011();

                                            th_hnx.Symbol = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                th_hnx.NumberofBids_QT = 0;

                                            }
                                            else { th_hnx.NumberofBids_QT = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                th_hnx.BidVolume_QT = 0;

                                            }
                                            else { th_hnx.BidVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                th_hnx.NumberofOffers_QT = 0;

                                            }
                                            else { th_hnx.NumberofOffers_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                th_hnx.OfferVolume_QT = 0;

                                            }
                                            else
                                            {
                                                th_hnx.OfferVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                th_hnx.Difference_QT = 0;

                                            }
                                            else { th_hnx.Difference_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                th_hnx.NumberofBids_NT = 0;

                                            }
                                            else { th_hnx.NumberofBids_NT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                            //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                th_hnx.BidVolume_NT = 0;

                                            }
                                            else { th_hnx.BidVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                th_hnx.NumberofOffers_NT = 0;

                                            }
                                            else { th_hnx.NumberofOffers_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                th_hnx.OfferVolume_NT = 0;

                                            }
                                            else { th_hnx.OfferVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                th_hnx.Difference_NT = 0;

                                            }
                                            else { th_hnx.Difference_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                th_hnx.SLDatMua = 0;

                                            }
                                            else { th_hnx.SLDatMua = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                th_hnx.KLDatMua = 0;

                                            }
                                            else { th_hnx.KLDatMua = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                th_hnx.SLDatBan = 0;

                                            }
                                            else { th_hnx.SLDatBan = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                th_hnx.KLDatBan = 0;

                                            }
                                            else { th_hnx.KLDatBan = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            th_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, th_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TH_DATLENH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //NDTNN_2011
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.NDTNN_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenNDTNN_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.NDTNN_2011.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                        {
                                            NDTNN2011 nt_hnx = new NDTNN2011();
                                            //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,

                                            nt_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            nt_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                nt_hnx.KLCKMAX = 0;

                                            }
                                            else { nt_hnx.KLCKMAX = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                nt_hnx.KLMUA_QT = 0;

                                            }
                                            else { nt_hnx.KLMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                nt_hnx.GTMUA_QT = 0;

                                            }
                                            else { nt_hnx.GTMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                nt_hnx.KLBAN_QT = 0;

                                            }
                                            else
                                            {
                                                nt_hnx.KLBAN_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                nt_hnx.GIATRI_QT = 0;

                                            }
                                            else { nt_hnx.GIATRI_QT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                nt_hnx.KLMUA_NT = 0;

                                            }
                                            else { nt_hnx.KLMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                            //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                nt_hnx.GTMUA_NT = 0;

                                            }
                                            else { nt_hnx.GTMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            //,KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                nt_hnx.KLBAN_NT = 0;

                                            }
                                            else { nt_hnx.KLBAN_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                nt_hnx.GIATRI_NT = 0;

                                            }
                                            else { nt_hnx.GIATRI_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                nt_hnx.CurrentRoom = 0;

                                            }
                                            else { nt_hnx.CurrentRoom = Convert.ToDouble(dataTable.Rows[i][column[11]]); }

                                            nt_hnx.KLLH = 0;
                                            nt_hnx.NamGiuMax = 0;

                                            nt_hnx.KLNDTN = 0;

                                            nt_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, nt_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.NDTNN_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //KQGD_2011 KQGD chi tiet
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGD_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenKQGD_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.KQGD_2011.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Q2"].ToString() != "")
                                        {
                                            KQGDCHITIET2011 ct_hnx = new KQGDCHITIET2011();
                                            //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighPrice,

                                            ct_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            ct_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                ct_hnx.BasicPrice = 0;

                                            }
                                            else { ct_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                ct_hnx.OpenPrice = 0;

                                            }
                                            else { ct_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                ct_hnx.ClosePrice = 0;

                                            }
                                            else { ct_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                ct_hnx.HighPrice = 0;

                                            }
                                            else
                                            {
                                                ct_hnx.HighPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            // //LowPrice,AveragePrice,NetChange,

                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                ct_hnx.LowPrice = 0;

                                            }
                                            else { ct_hnx.LowPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                ct_hnx.AveragePrice = 0;

                                            }
                                            else { ct_hnx.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                ct_hnx.NetChange = 0;

                                            }
                                            else { ct_hnx.NetChange = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                            //,Value_TC,GiaTriTT,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_BG = 0;

                                            }
                                            else { ct_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                ct_hnx.Value_BG = 0;

                                            }
                                            else { ct_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                ct_hnx.AveragePrice_TT = 0;

                                            }
                                            else { ct_hnx.AveragePrice_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_TT = 0;

                                            }
                                            else { ct_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                ct_hnx.Value_TT = 0;

                                            }
                                            else { ct_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_TC = 0;

                                            }
                                            else { ct_hnx.Volume_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                            {
                                                ct_hnx.Value_TC = 0;

                                            }
                                            else { ct_hnx.Value_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                            if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                            {
                                                ct_hnx.GiaTriTT = 0;

                                            }
                                            else { ct_hnx.GiaTriTT = Convert.ToDouble(dataTable.Rows[i][column[16]]); }


                                            ct_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, ct_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //KQGDTH_2011 KQGDTH
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGDTH_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenKQGDTH_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.KQGDTH_2011.BeginCell.Split(',');
                                    for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                    {

                                        KQGDTH2011 KQ_hnx = new KQGDTH2011();
                                        //
                                        //TypeName,Volume_BG,Value_BG,Weight_BG,Volume_TT,Value_TT,

                                        KQ_hnx.TypeName = dataTable.Rows[i][column[0]].ToString();
                                        if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_BG = 0;

                                        }
                                        else { KQ_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_BG = 0;

                                        }
                                        else { KQ_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_BG = 0;

                                        }
                                        else { KQ_hnx.Weight_BG = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                        //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_TT = 0;

                                        }
                                        else
                                        {
                                            KQ_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                        }
                                        //  //Weight_TT,Volume_MT,Value_MT,Weight_MT,Trangding_Date


                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_TT = 0;

                                        }
                                        else { KQ_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_TT = 0;

                                        }
                                        else { KQ_hnx.Weight_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_MT = 0;

                                        }
                                        else { KQ_hnx.Volume_MT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                        //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                        //,Value_TC,GiaTriTT,Trangding_Date
                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_MT = 0;

                                        }
                                        else { KQ_hnx.Value_MT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_MT = 0;

                                        }
                                        else { KQ_hnx.Weight_MT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }

                                        KQ_hnx.Trangding_Date = dateFile;
                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, KQ_hnx, null, null, null, null, null, null, null, null, null, null, null);
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGDTH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //Top_2010
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;

                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.Top_2010.SheetName];


                                    //Top10CK_GTGDL

                                    DataTable dataTable = configTable.DatTenTop10CK_GTGDL_2010(dataSet);


                                    string[] column = configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Column3"].ToString() == "" && dataTable.Rows[i]["Column4"].ToString() != "")
                                        {
                                            Top10CK_GTGDL gtgdl = new Top10CK_GTGDL();
                                            //Symbol,ValueN,WeightN,Trangding_Date


                                            gtgdl.Symbol = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                gtgdl.ValueN = 0;

                                            }
                                            else { gtgdl.ValueN = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                gtgdl.WeightN = 0;

                                            }
                                            else { gtgdl.WeightN = Convert.ToDouble(dataTable.Rows[i][column[2]]); }


                                            gtgdl.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Top10CK_KLGDL
                                    DataTable dataTable1 = configTable.DatTenTop10CK_KLGDL_2010(dataSet);



                                    string[] column1 = configs.HNX_File_2010.Top_2010.Top10CK_KLGDL.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable1.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable1.Rows[i][column1[0]].ToString(), out view) && dataTable1.Rows[i][column1[1]].ToString() != "" && dataTable1.Rows[i][column1[0]].ToString() != "")
                                        {
                                            Top10CK_KLGDL klgdl = new Top10CK_KLGDL();
                                            //Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date


                                            klgdl.Symbol = dataTable1.Rows[i][column1[0]].ToString();

                                            klgdl.AvePrice = 0;


                                            if (!float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                            {
                                                klgdl.Volume = 0;

                                            }
                                            else { klgdl.Volume = Convert.ToDouble(dataTable1.Rows[i][column1[1]]); }

                                            klgdl.PhanTram = 0;


                                            if (!float.TryParse(dataTable1.Rows[i][column1[2]].ToString(), out view))
                                            {
                                                klgdl.WeightN = 0;

                                            }
                                            else { klgdl.WeightN = Convert.ToDouble(dataTable1.Rows[i][column1[2]]); }


                                            klgdl.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }

                                    //Top10CP_GTNYL
                                    DataTable dataTable2 = configTable.DatTenTop10CP_GTNYL_2010_1(dataSet);



                                    string[] column2 = configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable2.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable2.Rows[i][column2[0]].ToString(), out view) && dataTable2.Rows[i][column2[0]].ToString() != "" && float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                        {
                                            Top10CP_CLGMAX gtnyl = new Top10CP_CLGMAX();
                                            //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date


                                            gtnyl.Symbol = dataTable2.Rows[i][column2[0]].ToString();
                                            if (!float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                            {
                                                gtnyl.HighPrice = 0;

                                            }
                                            else { gtnyl.HighPrice = Convert.ToDouble(dataTable2.Rows[i][column2[1]]); }
                                            if (!float.TryParse(dataTable2.Rows[i][column2[2]].ToString(), out view))
                                            {
                                                gtnyl.LowPrice = 0;

                                            }
                                            else { gtnyl.LowPrice = Convert.ToDouble(dataTable2.Rows[i][column2[2]]); }

                                            if (!float.TryParse(dataTable2.Rows[i][column2[3]].ToString(), out view))
                                            {
                                                gtnyl.TyLeChenhLech = 0;

                                            }
                                            else { gtnyl.TyLeChenhLech = Convert.ToDouble(dataTable2.Rows[i][column2[3]]); }


                                            gtnyl.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, gtnyl, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }

                                    //Top10CK_TANGGIA
                                    DataTable dataTable3 = configTable.DatTenTop10CK_TANGGIA_2010(dataSet);

                                    string[] column3 = configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable3.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable3.Rows[i][column3[0]].ToString(), out view) && float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                        {
                                            Top10CK_TANGGIA2010 tg = new Top10CK_TANGGIA2010();
                                            //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date


                                            tg.Symbol = dataTable3.Rows[i][column3[0]].ToString();
                                            if (!float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                            {
                                                tg.AvePrice = 0;

                                            }
                                            else { tg.AvePrice = Convert.ToDouble(dataTable3.Rows[i][column3[1]]); }
                                            if (!float.TryParse(dataTable3.Rows[i][column3[2]].ToString(), out view))
                                            {
                                                tg.MucTang = 0;

                                            }
                                            else { tg.MucTang = Convert.ToDouble(dataTable3.Rows[i][column3[2]]); }

                                            if (!float.TryParse(dataTable3.Rows[i][column3[3]].ToString(), out view))
                                            {
                                                tg.TyLeTang = 0;

                                            }
                                            else { tg.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column3[3]]); }


                                            tg.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2010(tg, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //Trái phiếu
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.GD_TRAIPHIEU.SheetName];

                                    DataTable dataTable = configTable.DatTenTH_GDTRAIPHIEU_2010(dataSet);



                                    string[] column = configs.HNX_File_2010.GD_TRAIPHIEU.BeginCell.Split(',');
                                    for (int i = 1; i < dataTable.Rows.Count - 0; i++)
                                    {

                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                        {
                                            GD_TRAIPHIEU th_hnx = new GD_TRAIPHIEU();
                                            // STT,Symbol,KyHanNam,GiaGDDong,LaiSuat,LoiSuat,KLGD,GTGD,Trangding_Date
                                            th_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            th_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                th_hnx.KyHanNam = 0;

                                            }
                                            else { th_hnx.KyHanNam = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                th_hnx.GiaGDDong = 0;

                                            }
                                            else { th_hnx.GiaGDDong = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                th_hnx.LaiSuat = 0;

                                            }
                                            else { th_hnx.LaiSuat = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                th_hnx.LoiSuat = 0;

                                            }
                                            else
                                            {
                                                th_hnx.LoiSuat = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                th_hnx.KLGD = 0;

                                            }
                                            else { th_hnx.KLGD = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                th_hnx.GTGD = 0;

                                            }
                                            else { th_hnx.GTGD = Convert.ToDouble(dataTable.Rows[i][column[7]]); }

                                            th_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, null, th_hnx, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.GD_TRAIPHIEU.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                    }
                    else if (!namefile.Contains(miux2) && namefile.Contains("KQGD"))
                    {
                        if (dateFile > dateNew)
                        {
                            // DateTime dateTo = DateTime.ParseExact(configs.ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            try
                            {


                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];


                                        DataTable dataTable = configTable.DatTenTT_DKGD_2011_21(dataSet);


                                        //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                        string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {

                                                KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    dkgd_hnx.SLCP_DKGD = 0;

                                                }
                                                else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    dkgd_hnx.SLCP_LH = 0;

                                                }
                                                else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    dkgd_hnx.Co_Tuc_2010 = 0;

                                                }
                                                else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    dkgd_hnx.PE = 0;

                                                }
                                                else
                                                {
                                                    dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    dkgd_hnx.EPS2010 = 0;

                                                }
                                                else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }

                                                dkgd_hnx.KLGD_10PHIEN = 0;

                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    dkgd_hnx.ROE = 0;

                                                }
                                                else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    dkgd_hnx.ROA = 0;

                                                }
                                                else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    dkgd_hnx.BasicPrice_KT = 0;

                                                }
                                                else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    dkgd_hnx.CeilingPrice_KT = 0;

                                                }
                                                else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    dkgd_hnx.FloorPrice_KT = 0;

                                                }
                                                else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    dkgd_hnx.BinhQuan = 0;

                                                }
                                                else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }

                                                dkgd_hnx.Tong = 0;
                                                dkgd_hnx.Co_Tuc_2009 = 0;

                                                dkgd_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            //Tinh hinh dat lenh
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        // var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];
                                        DataTable dataSet;

                                        dataSet = reader.AsDataSet().Tables["Tinh hinh dt lenh"];
                                        if (dataSet == null)
                                        {
                                            dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];
                                        }
                                        DataTable dataTable = configTable.DatTenTH_DATLENH_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.TH_DATLENH_2011.BeginCell.Split(',');
                                        for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            string vi = dataTable.Rows[i][column[0]].ToString();
                                            if (dataTable.Rows[i][column[0]].ToString() != "" && vi.Length < 7)
                                            {
                                                TinhHinhDatLenh2011 th_hnx = new TinhHinhDatLenh2011();

                                                th_hnx.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofBids_QT = 0;

                                                }
                                                else { th_hnx.NumberofBids_QT = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    th_hnx.BidVolume_QT = 0;

                                                }
                                                else { th_hnx.BidVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofOffers_QT = 0;

                                                }
                                                else { th_hnx.NumberofOffers_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    th_hnx.OfferVolume_QT = 0;

                                                }
                                                else
                                                {
                                                    th_hnx.OfferVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    th_hnx.Difference_QT = 0;

                                                }
                                                else { th_hnx.Difference_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofBids_NT = 0;

                                                }
                                                else { th_hnx.NumberofBids_NT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    th_hnx.BidVolume_NT = 0;

                                                }
                                                else { th_hnx.BidVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofOffers_NT = 0;

                                                }
                                                else { th_hnx.NumberofOffers_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    th_hnx.OfferVolume_NT = 0;

                                                }
                                                else { th_hnx.OfferVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    th_hnx.Difference_NT = 0;

                                                }
                                                else { th_hnx.Difference_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    th_hnx.SLDatMua = 0;

                                                }
                                                else { th_hnx.SLDatMua = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    th_hnx.KLDatMua = 0;

                                                }
                                                else { th_hnx.KLDatMua = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    th_hnx.SLDatBan = 0;

                                                }
                                                else { th_hnx.SLDatBan = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    th_hnx.KLDatBan = 0;

                                                }
                                                else { th_hnx.KLDatBan = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                th_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, th_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TH_DATLENH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //NDTNN_2011
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.NDTNN_2011.SheetName];

                                        DataTable dataTable = configTable.DatTenNDTNN_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.NDTNN_2011.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NDTNN2011 nt_hnx = new NDTNN2011();
                                                //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,

                                                nt_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                nt_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    nt_hnx.KLCKMAX = 0;

                                                }
                                                else { nt_hnx.KLCKMAX = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    nt_hnx.KLMUA_QT = 0;

                                                }
                                                else { nt_hnx.KLMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    nt_hnx.GTMUA_QT = 0;

                                                }
                                                else { nt_hnx.GTMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    nt_hnx.KLBAN_QT = 0;

                                                }
                                                else
                                                {
                                                    nt_hnx.KLBAN_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    nt_hnx.GIATRI_QT = 0;

                                                }
                                                else { nt_hnx.GIATRI_QT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    nt_hnx.KLMUA_NT = 0;

                                                }
                                                else { nt_hnx.KLMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    nt_hnx.GTMUA_NT = 0;

                                                }
                                                else { nt_hnx.GTMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                //,KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    nt_hnx.KLBAN_NT = 0;

                                                }
                                                else { nt_hnx.KLBAN_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    nt_hnx.GIATRI_NT = 0;

                                                }
                                                else { nt_hnx.GIATRI_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    nt_hnx.CurrentRoom = 0;

                                                }
                                                else { nt_hnx.CurrentRoom = Convert.ToDouble(dataTable.Rows[i][column[11]]); }

                                                nt_hnx.KLLH = 0;
                                                nt_hnx.NamGiuMax = 0;

                                                nt_hnx.KLNDTN = 0;

                                                nt_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, nt_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.NDTNN_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //KQGD_2011 KQGD chi tiet
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGD_2011.SheetName];

                                        DataTable dataTable = configTable.DatTenKQGD_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.KQGD_2011.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Q2"].ToString() != "")
                                            {
                                                KQGDCHITIET2011 ct_hnx = new KQGDCHITIET2011();
                                                //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighPrice,

                                                ct_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                ct_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    ct_hnx.BasicPrice = 0;

                                                }
                                                else { ct_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    ct_hnx.OpenPrice = 0;

                                                }
                                                else { ct_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    ct_hnx.ClosePrice = 0;

                                                }
                                                else { ct_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    ct_hnx.HighPrice = 0;

                                                }
                                                else
                                                {
                                                    ct_hnx.HighPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                // //LowPrice,AveragePrice,NetChange,

                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    ct_hnx.LowPrice = 0;

                                                }
                                                else { ct_hnx.LowPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    ct_hnx.AveragePrice = 0;

                                                }
                                                else { ct_hnx.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    ct_hnx.NetChange = 0;

                                                }
                                                else { ct_hnx.NetChange = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                                //,Value_TC,GiaTriTT,Trangding_Date
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    ct_hnx.Volume_BG = 0;

                                                }
                                                else { ct_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    ct_hnx.Value_BG = 0;

                                                }
                                                else { ct_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    ct_hnx.AveragePrice_TT = 0;

                                                }
                                                else { ct_hnx.AveragePrice_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    ct_hnx.Volume_TT = 0;

                                                }
                                                else { ct_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    ct_hnx.Value_TT = 0;

                                                }
                                                else { ct_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    ct_hnx.Volume_TC = 0;

                                                }
                                                else { ct_hnx.Volume_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    ct_hnx.Value_TC = 0;

                                                }
                                                else { ct_hnx.Value_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    ct_hnx.GiaTriTT = 0;

                                                }
                                                else { ct_hnx.GiaTriTT = Convert.ToDouble(dataTable.Rows[i][column[16]]); }


                                                ct_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, ct_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //KQGDTH_2011 KQGDTH
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGDTH_2011.SheetName];

                                        DataTable dataTable = configTable.DatTenKQGDTH_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.KQGDTH_2011.BeginCell.Split(',');
                                        for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                        {

                                            KQGDTH2011 KQ_hnx = new KQGDTH2011();
                                            //
                                            //TypeName,Volume_BG,Value_BG,Weight_BG,Volume_TT,Value_TT,

                                            KQ_hnx.TypeName = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                KQ_hnx.Volume_BG = 0;

                                            }
                                            else { KQ_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                KQ_hnx.Value_BG = 0;

                                            }
                                            else { KQ_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                KQ_hnx.Weight_BG = 0;

                                            }
                                            else { KQ_hnx.Weight_BG = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                KQ_hnx.Volume_TT = 0;

                                            }
                                            else
                                            {
                                                KQ_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                            }
                                            //  //Weight_TT,Volume_MT,Value_MT,Weight_MT,Trangding_Date


                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                KQ_hnx.Value_TT = 0;

                                            }
                                            else { KQ_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                KQ_hnx.Weight_TT = 0;

                                            }
                                            else { KQ_hnx.Weight_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                KQ_hnx.Volume_MT = 0;

                                            }
                                            else { KQ_hnx.Volume_MT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                            //,Value_TC,GiaTriTT,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                KQ_hnx.Value_MT = 0;

                                            }
                                            else { KQ_hnx.Value_MT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                KQ_hnx.Weight_MT = 0;

                                            }
                                            else { KQ_hnx.Weight_MT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }

                                            KQ_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, KQ_hnx, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGDTH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //Top_2010
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;

                                        var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.Top_2010.SheetName];


                                        //Top10CK_GTGDL

                                        DataTable dataTable = configTable.DatTenTop10CK_GTGDL_2010(dataSet);


                                        string[] column = configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Column3"].ToString() == "" && dataTable.Rows[i]["Column4"].ToString() != "")
                                            {
                                                Top10CK_GTGDL gtgdl = new Top10CK_GTGDL();
                                                //Symbol,ValueN,WeightN,Trangding_Date


                                                gtgdl.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {
                                                    gtgdl.ValueN = 0;

                                                }
                                                else { gtgdl.ValueN = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    gtgdl.WeightN = 0;

                                                }
                                                else { gtgdl.WeightN = Convert.ToDouble(dataTable.Rows[i][column[2]]); }


                                                gtgdl.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Top10CK_KLGDL
                                        DataTable dataTable1 = configTable.DatTenTop10CK_KLGDL_2010(dataSet);



                                        string[] column1 = configs.HNX_File_2010.Top_2010.Top10CK_KLGDL.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable1.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable1.Rows[i][column1[0]].ToString(), out view) && dataTable1.Rows[i][column1[1]].ToString() != "" && dataTable1.Rows[i][column1[0]].ToString() != "")
                                            {
                                                Top10CK_KLGDL klgdl = new Top10CK_KLGDL();
                                                //Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date


                                                klgdl.Symbol = dataTable1.Rows[i][column1[0]].ToString();

                                                klgdl.AvePrice = 0;


                                                if (!float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                                {
                                                    klgdl.Volume = 0;

                                                }
                                                else { klgdl.Volume = Convert.ToDouble(dataTable1.Rows[i][column1[1]]); }

                                                klgdl.PhanTram = 0;


                                                if (!float.TryParse(dataTable1.Rows[i][column1[2]].ToString(), out view))
                                                {
                                                    klgdl.WeightN = 0;

                                                }
                                                else { klgdl.WeightN = Convert.ToDouble(dataTable1.Rows[i][column1[2]]); }


                                                klgdl.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        //Top10CP_GTNYL
                                        DataTable dataTable2 = configTable.DatTenTop10CP_GTNYL_2010_1(dataSet);



                                        string[] column2 = configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable2.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable2.Rows[i][column2[0]].ToString(), out view) && dataTable2.Rows[i][column2[0]].ToString() != "" && float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                            {
                                                Top10CP_CLGMAX gtnyl = new Top10CP_CLGMAX();
                                                //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date


                                                gtnyl.Symbol = dataTable2.Rows[i][column2[0]].ToString();
                                                if (!float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                                {
                                                    gtnyl.HighPrice = 0;

                                                }
                                                else { gtnyl.HighPrice = Convert.ToDouble(dataTable2.Rows[i][column2[1]]); }
                                                if (!float.TryParse(dataTable2.Rows[i][column2[2]].ToString(), out view))
                                                {
                                                    gtnyl.LowPrice = 0;

                                                }
                                                else { gtnyl.LowPrice = Convert.ToDouble(dataTable2.Rows[i][column2[2]]); }

                                                if (!float.TryParse(dataTable2.Rows[i][column2[3]].ToString(), out view))
                                                {
                                                    gtnyl.TyLeChenhLech = 0;

                                                }
                                                else { gtnyl.TyLeChenhLech = Convert.ToDouble(dataTable2.Rows[i][column2[3]]); }


                                                gtnyl.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, gtnyl, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        //Top10CK_TANGGIA
                                        DataTable dataTable3 = configTable.DatTenTop10CK_TANGGIA_2010(dataSet);

                                        string[] column3 = configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable3.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable3.Rows[i][column3[0]].ToString(), out view) && float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                            {
                                                Top10CK_TANGGIA2010 tg = new Top10CK_TANGGIA2010();
                                                //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date


                                                tg.Symbol = dataTable3.Rows[i][column3[0]].ToString();
                                                if (!float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                                {
                                                    tg.AvePrice = 0;

                                                }
                                                else { tg.AvePrice = Convert.ToDouble(dataTable3.Rows[i][column3[1]]); }
                                                if (!float.TryParse(dataTable3.Rows[i][column3[2]].ToString(), out view))
                                                {
                                                    tg.MucTang = 0;

                                                }
                                                else { tg.MucTang = Convert.ToDouble(dataTable3.Rows[i][column3[2]]); }

                                                if (!float.TryParse(dataTable3.Rows[i][column3[3]].ToString(), out view))
                                                {
                                                    tg.TyLeTang = 0;

                                                }
                                                else { tg.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column3[3]]); }


                                                tg.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2010(tg, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //Trái phiếu
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.GD_TRAIPHIEU.SheetName];

                                        DataTable dataTable = configTable.DatTenTH_GDTRAIPHIEU_2010(dataSet);



                                        string[] column = configs.HNX_File_2010.GD_TRAIPHIEU.BeginCell.Split(',');
                                        for (int i = 1; i < dataTable.Rows.Count - 0; i++)
                                        {

                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                GD_TRAIPHIEU th_hnx = new GD_TRAIPHIEU();
                                                // STT,Symbol,KyHanNam,GiaGDDong,LaiSuat,LoiSuat,KLGD,GTGD,Trangding_Date
                                                th_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                th_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    th_hnx.KyHanNam = 0;

                                                }
                                                else { th_hnx.KyHanNam = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    th_hnx.GiaGDDong = 0;

                                                }
                                                else { th_hnx.GiaGDDong = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    th_hnx.LaiSuat = 0;

                                                }
                                                else { th_hnx.LaiSuat = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    th_hnx.LoiSuat = 0;

                                                }
                                                else
                                                {
                                                    th_hnx.LoiSuat = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    th_hnx.KLGD = 0;

                                                }
                                                else { th_hnx.KLGD = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    th_hnx.GTGD = 0;

                                                }
                                                else { th_hnx.GTGD = Convert.ToDouble(dataTable.Rows[i][column[7]]); }

                                                th_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, null, th_hnx, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.GD_TRAIPHIEU.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                        }
                        else
                        {
                            // DateTime dateTo = DateTime.ParseExact(configs.ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            try
                            {


                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables["Sheet5"];


                                        DataTable dataTable = configTable.DatTenTT_DKGD_2011_21(dataSet);


                                        //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                        string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {

                                                KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                                dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    dkgd_hnx.SLCP_DKGD = 0;

                                                }
                                                else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    dkgd_hnx.SLCP_LH = 0;

                                                }
                                                else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    dkgd_hnx.Co_Tuc_2010 = 0;

                                                }
                                                else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    dkgd_hnx.PE = 0;

                                                }
                                                else
                                                {
                                                    dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    dkgd_hnx.EPS2010 = 0;

                                                }
                                                else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }

                                                dkgd_hnx.KLGD_10PHIEN = 0;

                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    dkgd_hnx.ROE = 0;

                                                }
                                                else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    dkgd_hnx.ROA = 0;

                                                }
                                                else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    dkgd_hnx.BasicPrice_KT = 0;

                                                }
                                                else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    dkgd_hnx.CeilingPrice_KT = 0;

                                                }
                                                else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    dkgd_hnx.FloorPrice_KT = 0;

                                                }
                                                else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    dkgd_hnx.BinhQuan = 0;

                                                }
                                                else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }

                                                dkgd_hnx.Tong = 0;
                                                dkgd_hnx.Co_Tuc_2009 = 0;

                                                dkgd_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }
                            //Tinh hinh dat lenh
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        // var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];
                                        DataTable dataSet;

                                        dataSet = reader.AsDataSet().Tables["Sheet6"];

                                        DataTable dataTable = configTable.DatTenTH_DATLENH_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.TH_DATLENH_2011.BeginCell.Split(',');
                                        for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            string vi = dataTable.Rows[i][column[0]].ToString();
                                            if (dataTable.Rows[i][column[0]].ToString() != "" && vi.Length < 7)
                                            {
                                                TinhHinhDatLenh2011 th_hnx = new TinhHinhDatLenh2011();

                                                th_hnx.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofBids_QT = 0;

                                                }
                                                else { th_hnx.NumberofBids_QT = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    th_hnx.BidVolume_QT = 0;

                                                }
                                                else { th_hnx.BidVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofOffers_QT = 0;

                                                }
                                                else { th_hnx.NumberofOffers_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    th_hnx.OfferVolume_QT = 0;

                                                }
                                                else
                                                {
                                                    th_hnx.OfferVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    th_hnx.Difference_QT = 0;

                                                }
                                                else { th_hnx.Difference_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofBids_NT = 0;

                                                }
                                                else { th_hnx.NumberofBids_NT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    th_hnx.BidVolume_NT = 0;

                                                }
                                                else { th_hnx.BidVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    th_hnx.NumberofOffers_NT = 0;

                                                }
                                                else { th_hnx.NumberofOffers_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    th_hnx.OfferVolume_NT = 0;

                                                }
                                                else { th_hnx.OfferVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    th_hnx.Difference_NT = 0;

                                                }
                                                else { th_hnx.Difference_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    th_hnx.SLDatMua = 0;

                                                }
                                                else { th_hnx.SLDatMua = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    th_hnx.KLDatMua = 0;

                                                }
                                                else { th_hnx.KLDatMua = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    th_hnx.SLDatBan = 0;

                                                }
                                                else { th_hnx.SLDatBan = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    th_hnx.KLDatBan = 0;

                                                }
                                                else { th_hnx.KLDatBan = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                th_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, th_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }
                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TH_DATLENH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //NDTNN_2011
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables["Sheet4"];

                                        DataTable dataTable = configTable.DatTenNDTNN_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.NDTNN_2011.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                            {
                                                NDTNN2011 nt_hnx = new NDTNN2011();
                                                //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,

                                                nt_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                nt_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    nt_hnx.KLCKMAX = 0;

                                                }
                                                else { nt_hnx.KLCKMAX = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    nt_hnx.KLMUA_QT = 0;

                                                }
                                                else { nt_hnx.KLMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    nt_hnx.GTMUA_QT = 0;

                                                }
                                                else { nt_hnx.GTMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    nt_hnx.KLBAN_QT = 0;

                                                }
                                                else
                                                {
                                                    nt_hnx.KLBAN_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    nt_hnx.GIATRI_QT = 0;

                                                }
                                                else { nt_hnx.GIATRI_QT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    nt_hnx.KLMUA_NT = 0;

                                                }
                                                else { nt_hnx.KLMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                                //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    nt_hnx.GTMUA_NT = 0;

                                                }
                                                else { nt_hnx.GTMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                //,KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    nt_hnx.KLBAN_NT = 0;

                                                }
                                                else { nt_hnx.KLBAN_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    nt_hnx.GIATRI_NT = 0;

                                                }
                                                else { nt_hnx.GIATRI_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    nt_hnx.CurrentRoom = 0;

                                                }
                                                else { nt_hnx.CurrentRoom = Convert.ToDouble(dataTable.Rows[i][column[11]]); }

                                                nt_hnx.KLLH = 0;
                                                nt_hnx.NamGiuMax = 0;

                                                nt_hnx.KLNDTN = 0;

                                                nt_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, nt_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.NDTNN_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //KQGD_2011 KQGD chi tiet
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables["Sheet2"];

                                        DataTable dataTable = configTable.DatTenKQGD_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.KQGD_2011.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Q2"].ToString() != "")
                                            {
                                                KQGDCHITIET2011 ct_hnx = new KQGDCHITIET2011();
                                                //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighPrice,

                                                ct_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                ct_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    ct_hnx.BasicPrice = 0;

                                                }
                                                else { ct_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    ct_hnx.OpenPrice = 0;

                                                }
                                                else { ct_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    ct_hnx.ClosePrice = 0;

                                                }
                                                else { ct_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    ct_hnx.HighPrice = 0;

                                                }
                                                else
                                                {
                                                    ct_hnx.HighPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                }
                                                // //LowPrice,AveragePrice,NetChange,

                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    ct_hnx.LowPrice = 0;

                                                }
                                                else { ct_hnx.LowPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    ct_hnx.AveragePrice = 0;

                                                }
                                                else { ct_hnx.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    ct_hnx.NetChange = 0;

                                                }
                                                else { ct_hnx.NetChange = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                                //,Value_TC,GiaTriTT,Trangding_Date
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    ct_hnx.Volume_BG = 0;

                                                }
                                                else { ct_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    ct_hnx.Value_BG = 0;

                                                }
                                                else { ct_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    ct_hnx.AveragePrice_TT = 0;

                                                }
                                                else { ct_hnx.AveragePrice_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    ct_hnx.Volume_TT = 0;

                                                }
                                                else { ct_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    ct_hnx.Value_TT = 0;

                                                }
                                                else { ct_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    ct_hnx.Volume_TC = 0;

                                                }
                                                else { ct_hnx.Volume_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    ct_hnx.Value_TC = 0;

                                                }
                                                else { ct_hnx.Value_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                                if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                                {
                                                    ct_hnx.GiaTriTT = 0;

                                                }
                                                else { ct_hnx.GiaTriTT = Convert.ToDouble(dataTable.Rows[i][column[16]]); }


                                                ct_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, ct_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //KQGDTH_2011 KQGDTH
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        var dataSet = reader.AsDataSet().Tables["Sheet1"];

                                        DataTable dataTable = configTable.DatTenKQGDTH_2011(dataSet);



                                        string[] column = configs.HNX_File_2011.KQGDTH_2011.BeginCell.Split(',');
                                        for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                        {

                                            KQGDTH2011 KQ_hnx = new KQGDTH2011();
                                            //
                                            //TypeName,Volume_BG,Value_BG,Weight_BG,Volume_TT,Value_TT,

                                            KQ_hnx.TypeName = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                KQ_hnx.Volume_BG = 0;

                                            }
                                            else { KQ_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                KQ_hnx.Value_BG = 0;

                                            }
                                            else { KQ_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                KQ_hnx.Weight_BG = 0;

                                            }
                                            else { KQ_hnx.Weight_BG = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                KQ_hnx.Volume_TT = 0;

                                            }
                                            else
                                            {
                                                KQ_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                            }
                                            //  //Weight_TT,Volume_MT,Value_MT,Weight_MT,Trangding_Date


                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                KQ_hnx.Value_TT = 0;

                                            }
                                            else { KQ_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                KQ_hnx.Weight_TT = 0;

                                            }
                                            else { KQ_hnx.Weight_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                KQ_hnx.Volume_MT = 0;

                                            }
                                            else { KQ_hnx.Volume_MT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                            //,Value_TC,GiaTriTT,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                KQ_hnx.Value_MT = 0;

                                            }
                                            else { KQ_hnx.Value_MT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                KQ_hnx.Weight_MT = 0;

                                            }
                                            else { KQ_hnx.Weight_MT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }

                                            KQ_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, KQ_hnx, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGDTH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Console.WriteLine("File: " + filePath);



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //Top_2010
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        DataTable dataSet;
                                        dataSet = reader.AsDataSet().Tables["Sheet7"];

                                        if (dataSet == null)
                                        {
                                            dataSet = reader.AsDataSet().Tables["Sheet3"];
                                        }
                                        else
                                        {
                                            string views = dataSet.Rows[1]["Column0"].ToString();
                                            if (views.Contains("STT"))
                                            {
                                                dataSet = reader.AsDataSet().Tables["Sheet3"];
                                            }

                                        }


                                        //Top10CK_GTGDL

                                        DataTable dataTable = configTable.DatTenTop10CK_GTGDL_2010(dataSet);


                                        string[] column = configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Column3"].ToString() == "" && dataTable.Rows[i]["Column4"].ToString() != "")
                                            {
                                                Top10CK_GTGDL gtgdl = new Top10CK_GTGDL();
                                                //Symbol,ValueN,WeightN,Trangding_Date


                                                gtgdl.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {
                                                    gtgdl.ValueN = 0;

                                                }
                                                else { gtgdl.ValueN = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    gtgdl.WeightN = 0;

                                                }
                                                else { gtgdl.WeightN = Convert.ToDouble(dataTable.Rows[i][column[2]]); }


                                                gtgdl.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }
                                        //Top10CK_KLGDL
                                        DataTable dataTable1 = configTable.DatTenTop10CK_KLGDL_2010(dataSet);



                                        string[] column1 = configs.HNX_File_2010.Top_2010.Top10CK_KLGDL.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable1.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable1.Rows[i][column1[0]].ToString(), out view) && dataTable1.Rows[i][column1[1]].ToString() != "" && dataTable1.Rows[i][column1[0]].ToString() != "")
                                            {
                                                Top10CK_KLGDL klgdl = new Top10CK_KLGDL();
                                                //Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date


                                                klgdl.Symbol = dataTable1.Rows[i][column1[0]].ToString();

                                                klgdl.AvePrice = 0;


                                                if (!float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                                {
                                                    klgdl.Volume = 0;

                                                }
                                                else { klgdl.Volume = Convert.ToDouble(dataTable1.Rows[i][column1[1]]); }

                                                klgdl.PhanTram = 0;


                                                if (!float.TryParse(dataTable1.Rows[i][column1[2]].ToString(), out view))
                                                {
                                                    klgdl.WeightN = 0;

                                                }
                                                else { klgdl.WeightN = Convert.ToDouble(dataTable1.Rows[i][column1[2]]); }


                                                klgdl.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        //Top10CP_GTNYL
                                        DataTable dataTable2 = configTable.DatTenTop10CP_GTNYL_2010_1(dataSet);



                                        string[] column2 = configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable2.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable2.Rows[i][column2[0]].ToString(), out view) && dataTable2.Rows[i][column2[0]].ToString() != "" && float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                            {
                                                Top10CP_CLGMAX gtnyl = new Top10CP_CLGMAX();
                                                //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date


                                                gtnyl.Symbol = dataTable2.Rows[i][column2[0]].ToString();
                                                if (!float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                                {
                                                    gtnyl.HighPrice = 0;

                                                }
                                                else { gtnyl.HighPrice = Convert.ToDouble(dataTable2.Rows[i][column2[1]]); }
                                                if (!float.TryParse(dataTable2.Rows[i][column2[2]].ToString(), out view))
                                                {
                                                    gtnyl.LowPrice = 0;

                                                }
                                                else { gtnyl.LowPrice = Convert.ToDouble(dataTable2.Rows[i][column2[2]]); }

                                                if (!float.TryParse(dataTable2.Rows[i][column2[3]].ToString(), out view))
                                                {
                                                    gtnyl.TyLeChenhLech = 0;

                                                }
                                                else { gtnyl.TyLeChenhLech = Convert.ToDouble(dataTable2.Rows[i][column2[3]]); }


                                                gtnyl.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, gtnyl, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                        //Top10CK_TANGGIA
                                        DataTable dataTable3 = configTable.DatTenTop10CK_TANGGIA_2010(dataSet);

                                        string[] column3 = configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.BeginCell.Split(',');
                                        for (int i = 2; i < dataTable3.Rows.Count - 0; i++)
                                        {
                                            if (!float.TryParse(dataTable3.Rows[i][column3[0]].ToString(), out view) && float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                            {
                                                Top10CK_TANGGIA2010 tg = new Top10CK_TANGGIA2010();
                                                //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date


                                                tg.Symbol = dataTable3.Rows[i][column3[0]].ToString();
                                                if (!float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                                {
                                                    tg.AvePrice = 0;

                                                }
                                                else { tg.AvePrice = Convert.ToDouble(dataTable3.Rows[i][column3[1]]); }
                                                if (!float.TryParse(dataTable3.Rows[i][column3[2]].ToString(), out view))
                                                {
                                                    tg.MucTang = 0;

                                                }
                                                else { tg.MucTang = Convert.ToDouble(dataTable3.Rows[i][column3[2]]); }

                                                if (!float.TryParse(dataTable3.Rows[i][column3[3]].ToString(), out view))
                                                {
                                                    tg.TyLeTang = 0;

                                                }
                                                else { tg.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column3[3]]); }


                                                tg.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2010(tg, null, null, null);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }

                                        }

                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }



                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                            //Trái phiếu
                            try
                            {

                                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                                    {
                                        EBulkScript eBulkScript = new EBulkScript();
                                        float view;
                                        DataTable dataSet;
                                        DataTable dataSet2;
                                        dataSet = reader.AsDataSet().Tables["Sheet3"];
                                        dataSet2 = reader.AsDataSet().Tables["Sheet7"];
                                        if (dataSet.Select().Length <= 0)
                                        {
                                            dataSet = reader.AsDataSet().Tables["Sheet8"];
                                        }
                                        if (dataSet2 != null)
                                        {
                                            string views = dataSet.Rows[1]["Column0"].ToString();
                                            if (views.Contains("STT"))
                                            {
                                                dataSet2 = reader.AsDataSet().Tables["Sheet3"];
                                            }

                                        }

                                        if (dataSet2 != null)
                                        {
                                            DataTable dataTable = configTable.DatTenTH_GDTRAIPHIEU_2010(dataSet);



                                            string[] column = configs.HNX_File_2010.GD_TRAIPHIEU.BeginCell.Split(',');
                                            for (int i = 1; i < dataTable.Rows.Count - 0; i++)
                                            {

                                                if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    GD_TRAIPHIEU th_hnx = new GD_TRAIPHIEU();
                                                    // STT,Symbol,KyHanNam,GiaGDDong,LaiSuat,LoiSuat,KLGD,GTGD,Trangding_Date
                                                    th_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                                    th_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                                    if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                    {
                                                        th_hnx.KyHanNam = 0;

                                                    }
                                                    else { th_hnx.KyHanNam = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                    {
                                                        th_hnx.GiaGDDong = 0;

                                                    }
                                                    else { th_hnx.GiaGDDong = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                    {
                                                        th_hnx.LaiSuat = 0;

                                                    }
                                                    else { th_hnx.LaiSuat = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                    {
                                                        th_hnx.LoiSuat = 0;

                                                    }
                                                    else
                                                    {
                                                        th_hnx.LoiSuat = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                                    }
                                                    if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                    {
                                                        th_hnx.KLGD = 0;

                                                    }
                                                    else { th_hnx.KLGD = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                    if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                    {
                                                        th_hnx.GTGD = 0;

                                                    }
                                                    else { th_hnx.GTGD = Convert.ToDouble(dataTable.Rows[i][column[7]]); }

                                                    th_hnx.Trangding_Date = dateFile;
                                                    eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, null, th_hnx, null);
                                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                        mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                    // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                                }
                                            }
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            {
                                                // exec script mssql+oracle
                                                string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.GD_TRAIPHIEU.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                                configTable.ExecBulkScript(test);
                                                mssqlBuilder_HNX.Clear();

                                            }
                                            //Console.WriteLine("File: " + filePath);


                                        }


                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("File erorr: " + filePath);
                            }

                            //============================================================================//
                        }

                    }
                    else
                    {
                        // DateTime dateTo = DateTime.ParseExact(configs.ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        try
                        {


                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];


                                    DataTable dataTable = configTable.DatTenTT_DKGD_2011_21(dataSet);


                                    //    DataTable dataTable = dataSetX.Tables[configs.HNX_File_2011.TT_DKGD_2011.SheetName];

                                    string[] column = configs.HNX_File_2011.TT_DKGD_2011.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 1; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                        {

                                            KQGIAODICHCP2011 dkgd_hnx = new KQGIAODICHCP2011();

                                            dkgd_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            dkgd_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                dkgd_hnx.SLCP_DKGD = 0;

                                            }
                                            else { dkgd_hnx.SLCP_DKGD = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                dkgd_hnx.SLCP_LH = 0;

                                            }
                                            else { dkgd_hnx.SLCP_LH = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                dkgd_hnx.Co_Tuc_2010 = 0;

                                            }
                                            else { dkgd_hnx.Co_Tuc_2010 = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                dkgd_hnx.PE = 0;

                                            }
                                            else
                                            {
                                                dkgd_hnx.PE = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                dkgd_hnx.EPS2010 = 0;

                                            }
                                            else { dkgd_hnx.EPS2010 = Convert.ToDouble(dataTable.Rows[i][column[6]]); }

                                            dkgd_hnx.KLGD_10PHIEN = 0;

                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                dkgd_hnx.ROE = 0;

                                            }
                                            else { dkgd_hnx.ROE = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                dkgd_hnx.ROA = 0;

                                            }
                                            else { dkgd_hnx.ROA = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                dkgd_hnx.BasicPrice_KT = 0;

                                            }
                                            else { dkgd_hnx.BasicPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                dkgd_hnx.CeilingPrice_KT = 0;

                                            }
                                            else { dkgd_hnx.CeilingPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                dkgd_hnx.FloorPrice_KT = 0;

                                            }
                                            else { dkgd_hnx.FloorPrice_KT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                dkgd_hnx.BinhQuan = 0;

                                            }
                                            else { dkgd_hnx.BinhQuan = Convert.ToDouble(dataTable.Rows[i][column[12]]); }

                                            dkgd_hnx.Tong = 0;
                                            dkgd_hnx.Co_Tuc_2009 = 0;

                                            dkgd_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TT_DKGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }
                        //Tinh hinh dat lenh
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    // var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];
                                    DataTable dataSet;

                                    dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.TH_DATLENH_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenTH_DATLENH_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.TH_DATLENH_2011.BeginCell.Split(',');
                                    for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        string vi = dataTable.Rows[i][column[0]].ToString();
                                        if (dataTable.Rows[i][column[0]].ToString() != "" && vi.Length < 7)
                                        {
                                            TinhHinhDatLenh2011 th_hnx = new TinhHinhDatLenh2011();

                                            th_hnx.Symbol = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                th_hnx.NumberofBids_QT = 0;

                                            }
                                            else { th_hnx.NumberofBids_QT = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                th_hnx.BidVolume_QT = 0;

                                            }
                                            else { th_hnx.BidVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                th_hnx.NumberofOffers_QT = 0;

                                            }
                                            else { th_hnx.NumberofOffers_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                th_hnx.OfferVolume_QT = 0;

                                            }
                                            else
                                            {
                                                th_hnx.OfferVolume_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                th_hnx.Difference_QT = 0;

                                            }
                                            else { th_hnx.Difference_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                th_hnx.NumberofBids_NT = 0;

                                            }
                                            else { th_hnx.NumberofBids_NT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                            //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                th_hnx.BidVolume_NT = 0;

                                            }
                                            else { th_hnx.BidVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                th_hnx.NumberofOffers_NT = 0;

                                            }
                                            else { th_hnx.NumberofOffers_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                th_hnx.OfferVolume_NT = 0;

                                            }
                                            else { th_hnx.OfferVolume_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                th_hnx.Difference_NT = 0;

                                            }
                                            else { th_hnx.Difference_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                th_hnx.SLDatMua = 0;

                                            }
                                            else { th_hnx.SLDatMua = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                th_hnx.KLDatMua = 0;

                                            }
                                            else { th_hnx.KLDatMua = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                th_hnx.SLDatBan = 0;

                                            }
                                            else { th_hnx.SLDatBan = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                th_hnx.KLDatBan = 0;

                                            }
                                            else { th_hnx.KLDatBan = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            th_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, th_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.TH_DATLENH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //NDTNN_2011
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.NDTNN_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenNDTNN_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.NDTNN_2011.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                        {
                                            NDTNN2011 nt_hnx = new NDTNN2011();
                                            //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,

                                            nt_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            nt_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                nt_hnx.KLCKMAX = 0;

                                            }
                                            else { nt_hnx.KLCKMAX = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                nt_hnx.KLMUA_QT = 0;

                                            }
                                            else { nt_hnx.KLMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                nt_hnx.GTMUA_QT = 0;

                                            }
                                            else { nt_hnx.GTMUA_QT = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                nt_hnx.KLBAN_QT = 0;

                                            }
                                            else
                                            {
                                                nt_hnx.KLBAN_QT = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                nt_hnx.GIATRI_QT = 0;

                                            }
                                            else { nt_hnx.GIATRI_QT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                nt_hnx.KLMUA_NT = 0;

                                            }
                                            else { nt_hnx.KLMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                                            //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date 
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                nt_hnx.GTMUA_NT = 0;

                                            }
                                            else { nt_hnx.GTMUA_NT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            //,KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                nt_hnx.KLBAN_NT = 0;

                                            }
                                            else { nt_hnx.KLBAN_NT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                nt_hnx.GIATRI_NT = 0;

                                            }
                                            else { nt_hnx.GIATRI_NT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                nt_hnx.CurrentRoom = 0;

                                            }
                                            else { nt_hnx.CurrentRoom = Convert.ToDouble(dataTable.Rows[i][column[11]]); }

                                            nt_hnx.KLLH = 0;
                                            nt_hnx.NamGiuMax = 0;

                                            nt_hnx.KLNDTN = 0;

                                            nt_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, nt_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.NDTNN_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //KQGD_2011 KQGD chi tiet
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGD_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenKQGD_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.KQGD_2011.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view))
                                        {
                                            KQGDCHITIET2011 ct_hnx = new KQGDCHITIET2011();
                                            //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighPrice,

                                            ct_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            ct_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                ct_hnx.BasicPrice = 0;

                                            }
                                            else { ct_hnx.BasicPrice = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                ct_hnx.OpenPrice = 0;

                                            }
                                            else { ct_hnx.OpenPrice = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                ct_hnx.ClosePrice = 0;

                                            }
                                            else { ct_hnx.ClosePrice = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                ct_hnx.HighPrice = 0;

                                            }
                                            else
                                            {
                                                ct_hnx.HighPrice = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            // //LowPrice,AveragePrice,NetChange,

                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                ct_hnx.LowPrice = 0;

                                            }
                                            else { ct_hnx.LowPrice = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                ct_hnx.AveragePrice = 0;

                                            }
                                            else { ct_hnx.AveragePrice = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                            {
                                                ct_hnx.NetChange = 0;

                                            }
                                            else { ct_hnx.NetChange = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                            //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                            //,Value_TC,GiaTriTT,Trangding_Date
                                            if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_BG = 0;

                                            }
                                            else { ct_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[9]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                            {
                                                ct_hnx.Value_BG = 0;

                                            }
                                            else { ct_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                            {
                                                ct_hnx.AveragePrice_TT = 0;

                                            }
                                            else { ct_hnx.AveragePrice_TT = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_TT = 0;

                                            }
                                            else { ct_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                            {
                                                ct_hnx.Value_TT = 0;

                                            }
                                            else { ct_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                            {
                                                ct_hnx.Volume_TC = 0;

                                            }
                                            else { ct_hnx.Volume_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                            {
                                                ct_hnx.Value_TC = 0;

                                            }
                                            else { ct_hnx.Value_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                            if (!float.TryParse(dataTable.Rows[i][column[16]].ToString(), out view))
                                            {
                                                ct_hnx.GiaTriTT = 0;

                                            }
                                            else { ct_hnx.GiaTriTT = Convert.ToDouble(dataTable.Rows[i][column[16]]); }


                                            ct_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, ct_hnx, null, null, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGD_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //KQGDTH_2011 KQGDTH
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2011.KQGDTH_2011.SheetName];

                                    DataTable dataTable = configTable.DatTenKQGDTH_2011(dataSet);



                                    string[] column = configs.HNX_File_2011.KQGDTH_2011.BeginCell.Split(',');
                                    for (int i = 6; i < dataTable.Rows.Count - 0; i++)
                                    {

                                        KQGDTH2011 KQ_hnx = new KQGDTH2011();
                                        //
                                        //TypeName,Volume_BG,Value_BG,Weight_BG,Volume_TT,Value_TT,

                                        KQ_hnx.TypeName = dataTable.Rows[i][column[0]].ToString();
                                        if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_BG = 0;

                                        }
                                        else { KQ_hnx.Volume_BG = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_BG = 0;

                                        }
                                        else { KQ_hnx.Value_BG = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_BG = 0;

                                        }
                                        else { KQ_hnx.Weight_BG = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                        //KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT

                                        if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_TT = 0;

                                        }
                                        else
                                        {
                                            KQ_hnx.Volume_TT = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                        }
                                        //  //Weight_TT,Volume_MT,Value_MT,Weight_MT,Trangding_Date


                                        if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_TT = 0;

                                        }
                                        else { KQ_hnx.Value_TT = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_TT = 0;

                                        }
                                        else { KQ_hnx.Weight_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                        {
                                            KQ_hnx.Volume_MT = 0;

                                        }
                                        else { KQ_hnx.Volume_MT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                        //Volume_BG,Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC
                                        //,Value_TC,GiaTriTT,Trangding_Date
                                        if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                        {
                                            KQ_hnx.Value_MT = 0;

                                        }
                                        else { KQ_hnx.Value_MT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                        if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                        {
                                            KQ_hnx.Weight_MT = 0;

                                        }
                                        else { KQ_hnx.Weight_MT = Convert.ToDouble(dataTable.Rows[i][column[9]]); }

                                        KQ_hnx.Trangding_Date = dateFile;
                                        eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, KQ_hnx, null, null, null, null, null, null, null, null, null, null, null);
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                            mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                        // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);



                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.KQGDTH_2011.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //Top_2010
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;

                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.Top_2010.SheetName];


                                    //Top10CK_GTGDL

                                    DataTable dataTable = configTable.DatTenTop10CK_GTGDL_2010(dataSet);


                                    string[] column = configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && dataTable.Rows[i]["Column3"].ToString() == "" && dataTable.Rows[i]["Column4"].ToString() != "")
                                        {
                                            Top10CK_GTGDL gtgdl = new Top10CK_GTGDL();
                                            //Symbol,ValueN,WeightN,Trangding_Date


                                            gtgdl.Symbol = dataTable.Rows[i][column[0]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                gtgdl.ValueN = 0;

                                            }
                                            else { gtgdl.ValueN = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                gtgdl.WeightN = 0;

                                            }
                                            else { gtgdl.WeightN = Convert.ToDouble(dataTable.Rows[i][column[2]]); }


                                            gtgdl.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GTGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Top10CK_KLGDL
                                    DataTable dataTable1 = configTable.DatTenTop10CK_KLGDL_2010(dataSet);



                                    string[] column1 = configs.HNX_File_2010.Top_2010.Top10CK_KLGDL.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable1.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable1.Rows[i][column1[0]].ToString(), out view) && dataTable1.Rows[i][column1[1]].ToString() != "" && dataTable1.Rows[i][column1[0]].ToString() != "")
                                        {
                                            Top10CK_KLGDL klgdl = new Top10CK_KLGDL();
                                            //Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date


                                            klgdl.Symbol = dataTable1.Rows[i][column1[0]].ToString();

                                            klgdl.AvePrice = 0;


                                            if (!float.TryParse(dataTable1.Rows[i][column1[1]].ToString(), out view))
                                            {
                                                klgdl.Volume = 0;

                                            }
                                            else { klgdl.Volume = Convert.ToDouble(dataTable1.Rows[i][column1[1]]); }

                                            klgdl.PhanTram = 0;


                                            if (!float.TryParse(dataTable1.Rows[i][column1[2]].ToString(), out view))
                                            {
                                                klgdl.WeightN = 0;

                                            }
                                            else { klgdl.WeightN = Convert.ToDouble(dataTable1.Rows[i][column1[2]]); }


                                            klgdl.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2011.Top_2011.Top10CK_KLGDL.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }

                                    //Top10CP_GTNYL
                                    DataTable dataTable2 = configTable.DatTenTop10CP_GTNYL_2010_1(dataSet);



                                    string[] column2 = configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable2.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable2.Rows[i][column2[0]].ToString(), out view) && dataTable2.Rows[i][column2[0]].ToString() != "" && float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                        {
                                            Top10CP_CLGMAX gtnyl = new Top10CP_CLGMAX();
                                            //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date


                                            gtnyl.Symbol = dataTable2.Rows[i][column2[0]].ToString();
                                            if (!float.TryParse(dataTable2.Rows[i][column2[1]].ToString(), out view))
                                            {
                                                gtnyl.HighPrice = 0;

                                            }
                                            else { gtnyl.HighPrice = Convert.ToDouble(dataTable2.Rows[i][column2[1]]); }
                                            if (!float.TryParse(dataTable2.Rows[i][column2[2]].ToString(), out view))
                                            {
                                                gtnyl.LowPrice = 0;

                                            }
                                            else { gtnyl.LowPrice = Convert.ToDouble(dataTable2.Rows[i][column2[2]]); }

                                            if (!float.TryParse(dataTable2.Rows[i][column2[3]].ToString(), out view))
                                            {
                                                gtnyl.TyLeChenhLech = 0;

                                            }
                                            else { gtnyl.TyLeChenhLech = Convert.ToDouble(dataTable2.Rows[i][column2[3]]); }


                                            gtnyl.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, gtnyl, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CP_CLGMAX.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }

                                    //Top10CK_TANGGIA
                                    DataTable dataTable3 = configTable.DatTenTop10CK_TANGGIA_2010(dataSet);

                                    string[] column3 = configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable3.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable3.Rows[i][column3[0]].ToString(), out view) && float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                        {
                                            Top10CK_TANGGIA2010 tg = new Top10CK_TANGGIA2010();
                                            //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date


                                            tg.Symbol = dataTable3.Rows[i][column3[0]].ToString();
                                            if (!float.TryParse(dataTable3.Rows[i][column3[1]].ToString(), out view))
                                            {
                                                tg.AvePrice = 0;

                                            }
                                            else { tg.AvePrice = Convert.ToDouble(dataTable3.Rows[i][column3[1]]); }
                                            if (!float.TryParse(dataTable3.Rows[i][column3[2]].ToString(), out view))
                                            {
                                                tg.MucTang = 0;

                                            }
                                            else { tg.MucTang = Convert.ToDouble(dataTable3.Rows[i][column3[2]]); }

                                            if (!float.TryParse(dataTable3.Rows[i][column3[3]].ToString(), out view))
                                            {
                                                tg.TyLeTang = 0;

                                            }
                                            else { tg.TyLeTang = Convert.ToDouble(dataTable3.Rows[i][column3[3]]); }


                                            tg.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2010(tg, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_TANGGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }

                                    //Top10CK_GIAMGIA
                                    DataTable dataTable4 = configTable.DatTenTop10CK_GIAMGIA_2010(dataSet);



                                    string[] column4 = configs.HNX_File_2010.Top_2010.Top10CK_GIAMGIA.BeginCell.Split(',');
                                    for (int i = 3; i < dataTable4.Rows.Count - 0; i++)
                                    {
                                        if (!float.TryParse(dataTable4.Rows[i][column4[0]].ToString(), out view) && dataTable4.Rows[i][column4[0]].ToString() != "" && float.TryParse(dataTable4.Rows[i][column4[1]].ToString(), out view))
                                        {
                                            Top10CK_GIAMGIA gg = new Top10CK_GIAMGIA();
                                            //Symbol,AvePrice,MucGiam,TyLeTang,Trangding_Date


                                            gg.Symbol = dataTable4.Rows[i][column4[0]].ToString();
                                            if (!float.TryParse(dataTable4.Rows[i][column4[1]].ToString(), out view))
                                            {
                                                gg.AvePrice = 0;

                                            }
                                            else { gg.AvePrice = Convert.ToDouble(dataTable4.Rows[i][column4[1]]); }
                                            if (!float.TryParse(dataTable4.Rows[i][column4[2]].ToString(), out view))
                                            {
                                                gg.MucGiam = 0;

                                            }
                                            else { gg.MucGiam = Convert.ToDouble(dataTable4.Rows[i][column4[2]]); }

                                            if (!float.TryParse(dataTable4.Rows[i][column4[3]].ToString(), out view))
                                            {
                                                gg.TyLeTang = 0;

                                            }
                                            else { gg.TyLeTang = Convert.ToDouble(dataTable4.Rows[i][column4[3]]); }


                                            gg.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2011(null, null, null, null, null, null, null, null, null, gg, null, null, null, null, null, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }

                                    }

                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.Top_2010.Top10CK_GIAMGIA.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }

                        //============================================================================//
                        //Trái phiếu
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.GD_TRAIPHIEU.SheetName];

                                    DataTable dataTable = configTable.DatTenTH_GDTRAIPHIEU_2010(dataSet);



                                    string[] column = configs.HNX_File_2010.GD_TRAIPHIEU.BeginCell.Split(',');
                                    for (int i = 2; i < dataTable.Rows.Count - 0; i++)
                                    {

                                        if (float.TryParse(dataTable.Rows[i][column[0]].ToString(), out view) && float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                        {
                                            GD_TRAIPHIEU th_hnx = new GD_TRAIPHIEU();
                                            // STT,Symbol,KyHanNam,GiaGDDong,LaiSuat,LoiSuat,KLGD,GTGD,Trangding_Date
                                            th_hnx.STT = Convert.ToInt32(dataTable.Rows[i][column[0]]);
                                            th_hnx.Symbol = dataTable.Rows[i][column[1]].ToString();
                                            if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                            {
                                                th_hnx.KyHanNam = 0;

                                            }
                                            else { th_hnx.KyHanNam = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                            {
                                                th_hnx.GiaGDDong = 0;

                                            }
                                            else { th_hnx.GiaGDDong = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                            {
                                                th_hnx.LaiSuat = 0;

                                            }
                                            else { th_hnx.LaiSuat = Convert.ToDouble(dataTable.Rows[i][column[4]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                            {
                                                th_hnx.LoiSuat = 0;

                                            }
                                            else
                                            {
                                                th_hnx.LoiSuat = Convert.ToDouble(dataTable.Rows[i][column[5]]);
                                            }
                                            if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                            {
                                                th_hnx.KLGD = 0;

                                            }
                                            else { th_hnx.KLGD = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                            if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                            {
                                                th_hnx.GTGD = 0;

                                            }
                                            else { th_hnx.GTGD = Convert.ToDouble(dataTable.Rows[i][column[7]]); }

                                            th_hnx.Trangding_Date = dateFile;
                                            eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, null, th_hnx, null);
                                            if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                            // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                        }
                                    }
                                    if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                    {
                                        // exec script mssql+oracle
                                        string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.GD_TRAIPHIEU.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                        configTable.ExecBulkScript(test);
                                        mssqlBuilder_HNX.Clear();

                                    }
                                    //Console.WriteLine("File: " + filePath);



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }
                        //============================================================================//

                        //GDTP NDTNN
                        try
                        {

                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    EBulkScript eBulkScript = new EBulkScript();
                                    float view;
                                    var dataSet = reader.AsDataSet().Tables[configs.HNX_File_2010.GDTP_NDTNN.SheetName];
                                    int rowTable = dataSet.Select().Length;
                                    if (rowTable > 0)
                                    {
                                        DataTable dataTable = configTable.DatTenTH_GDTP_NDTNN_2010(dataSet);



                                        string[] column = configs.HNX_File_2010.GDTP_NDTNN.BeginCell.Split(',');
                                        for (int i = 3; i < dataTable.Rows.Count - 1; i++)
                                        {
                                            if (float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                            {
                                                GDTP_NDTNN th_hnx = new GDTP_NDTNN();
                                                //Symbol,KLMua_KL,KLBan_KL,KL_ChenhLech,GTMua_KL,GTBan_KL,KLMua_TT
                                                //,KLBan_TT,KL_ChenhLech_TT,GTMua_TT,GTBan_TT
                                                //,KLMua_TC,KLBan_TC,KL_ChenhLech_TC,GTMua_TC,GTBan_TC,Trangding_Date

                                                th_hnx.Symbol = dataTable.Rows[i][column[0]].ToString();
                                                if (!float.TryParse(dataTable.Rows[i][column[1]].ToString(), out view))
                                                {
                                                    th_hnx.KLMua_KL = 0;

                                                }
                                                else { th_hnx.KLMua_KL = Convert.ToDouble(dataTable.Rows[i][column[1]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[2]].ToString(), out view))
                                                {
                                                    th_hnx.KLBan_KL = 0;

                                                }
                                                else { th_hnx.KLBan_KL = Convert.ToDouble(dataTable.Rows[i][column[2]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[3]].ToString(), out view))
                                                {
                                                    th_hnx.KL_ChenhLech = 0;

                                                }
                                                else { th_hnx.KL_ChenhLech = Convert.ToDouble(dataTable.Rows[i][column[3]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[4]].ToString(), out view))
                                                {
                                                    th_hnx.GTMua_KL = 0;

                                                }
                                                else
                                                {
                                                    th_hnx.GTMua_KL = Convert.ToDouble(dataTable.Rows[i][column[4]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[5]].ToString(), out view))
                                                {
                                                    th_hnx.GTBan_KL = 0;

                                                }
                                                else { th_hnx.GTBan_KL = Convert.ToDouble(dataTable.Rows[i][column[5]]); }
                                                //=============
                                                if (!float.TryParse(dataTable.Rows[i][column[6]].ToString(), out view))
                                                {
                                                    th_hnx.KLMua_TT = 0;

                                                }
                                                else { th_hnx.KLMua_TT = Convert.ToDouble(dataTable.Rows[i][column[6]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[7]].ToString(), out view))
                                                {
                                                    th_hnx.KLBan_TT = 0;

                                                }
                                                else { th_hnx.KLBan_TT = Convert.ToDouble(dataTable.Rows[i][column[7]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[8]].ToString(), out view))
                                                {
                                                    th_hnx.KL_ChenhLech_TT = 0;

                                                }
                                                else { th_hnx.KL_ChenhLech_TT = Convert.ToDouble(dataTable.Rows[i][column[8]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[9]].ToString(), out view))
                                                {
                                                    th_hnx.GTMua_TT = 0;

                                                }
                                                else
                                                {
                                                    th_hnx.GTMua_TT = Convert.ToDouble(dataTable.Rows[i][column[9]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[10]].ToString(), out view))
                                                {
                                                    th_hnx.GTBan_TT = 0;

                                                }
                                                else { th_hnx.GTBan_TT = Convert.ToDouble(dataTable.Rows[i][column[10]]); }
                                                //======
                                                if (!float.TryParse(dataTable.Rows[i][column[11]].ToString(), out view))
                                                {
                                                    th_hnx.KLMua_TC = 0;

                                                }
                                                else { th_hnx.KLMua_TC = Convert.ToDouble(dataTable.Rows[i][column[11]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[12]].ToString(), out view))
                                                {
                                                    th_hnx.KLBan_TC = 0;

                                                }
                                                else { th_hnx.KLBan_TC = Convert.ToDouble(dataTable.Rows[i][column[12]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[13]].ToString(), out view))
                                                {
                                                    th_hnx.KL_ChenhLech_TC = 0;

                                                }
                                                else { th_hnx.KL_ChenhLech_TC = Convert.ToDouble(dataTable.Rows[i][column[13]]); }
                                                if (!float.TryParse(dataTable.Rows[i][column[14]].ToString(), out view))
                                                {
                                                    th_hnx.GTMua_TC = 0;

                                                }
                                                else
                                                {
                                                    th_hnx.GTMua_TC = Convert.ToDouble(dataTable.Rows[i][column[14]]);
                                                }
                                                if (!float.TryParse(dataTable.Rows[i][column[15]].ToString(), out view))
                                                {
                                                    th_hnx.GTBan_TC = 0;

                                                }
                                                else { th_hnx.GTBan_TC = Convert.ToDouble(dataTable.Rows[i][column[15]]); }

                                                th_hnx.Trangding_Date = dateFile;
                                                eBulkScript = this.configTable.GetScriptTTCBHNX2010(null, null, null, th_hnx);
                                                if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                                    mssqlBuilder_HNX.Append(eBulkScript.MssqlScript);
                                                // eBulkScript = Update_TTICBVCTDKGD(ttcb, true);

                                            }


                                        }
                                        if (!string.IsNullOrEmpty(eBulkScript.MssqlScript))
                                        {
                                            // exec script mssql+oracle
                                            string test = EDalResult.__STRING_RETURN_NEW_LINE + EDalResult.__STRING_INSERT + configs.HNX_File_2010.GDTP_NDTNN.TableName + EDalResult.__STRING_VALUES + mssqlBuilder_HNX.ToString().TrimEnd(',') + ";";
                                            configTable.ExecBulkScript(test);
                                            mssqlBuilder_HNX.Clear();

                                        }

                                    }



                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("File erorr: " + filePath);
                        }
                        //============================================================================//
                    }


                }
            }

            //
        }

    }
}
