using ApiExcelToDB.Entities;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ApiExcelToDB.HOSE
{
    public class ReadExcelHose
    {
        public ReadExcelHose(string name)
        {
            Read_HSX(name);
        }
        SqlCommand command = new SqlCommand();
        private readonly ConfigTable configTable;

        public void Read_HSX(string name)
        {
            IConfiguration config = new ConfigurationBuilder()
                 .SetBasePath(Directory.GetCurrentDirectory())
                 .AddJsonFile("appsettingshsx.json")
                 .Build();
            try
            {
                Econfig configs = new Econfig();
                UpdateDB insert = new UpdateDB();
                string folderPath = config["FileFolder:folderPath"] + name;
                int bas = folderPath.IndexOf("KQGD Hose");
                string invalidFileListPath = config["FolderLogError:folderPathError"];

                using (StreamWriter writer = new StreamWriter(invalidFileListPath, true))
                {
                    var filePaths = Directory.GetFiles(folderPath, "*.xls", SearchOption.AllDirectories)
                     .OrderBy(f => f);

                    foreach (var filePath in filePaths)
                    {
                        try
                        {
                            string connectionString = config[configs.SQLConn];
                            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
                            {
                                sqlConnection.Open();
                                string fommatDate = configs.DATETIME_FORMAT_24;
                              //  string fommatDate1 = configs.DATETIME_FORMAT_23;
                                string regexRar = config[configs.regexRar];
                                string fike = filePath.Remove(1, bas - 1);
                                Match match = Regex.Match(fike, regexRar);


                                if (match.Success)
                                {
                                    string group6Value = match.Groups[4].Value;
                                    string group7Value = match.Groups[5].Value;

                                    DateTime group6Date;
                                    string[] dateFormats = { configs.DATETIME_FORMAT_21, configs.DATETIME_FORMAT_22 };
                                    if (DateTime.TryParseExact(group6Value, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out group6Date))
                                    {
                                        if (DateTime.TryParseExact(config[configs.fromDate1], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate1) && DateTime.TryParseExact(config[configs.toDate1], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate1) && DateTime.TryParseExact(config[configs.fromDate2], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate2)
                                                && DateTime.TryParseExact(config[configs.toDate2], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate2) && DateTime.TryParseExact(config[configs.fromDate3], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate3) && DateTime.TryParseExact(config[configs.toDate3], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate3)
                                                && DateTime.TryParseExact(config[configs.fromDate4], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate4) && DateTime.TryParseExact(config[configs.toDate4], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate4) && DateTime.TryParseExact(config[configs.fromDate5], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate5)
                                                && DateTime.TryParseExact(config[configs.toDate5], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate5) && DateTime.TryParseExact(config[configs.fromDate6], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate6) && DateTime.TryParseExact(config[configs.toDate6], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate6)
                                                && DateTime.TryParseExact(config[configs.fromDate7], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate7) && DateTime.TryParseExact(config[configs.toDate7], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate7) && DateTime.TryParseExact(config[configs.fromDate8], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate8)
                                                && DateTime.TryParseExact(config[configs.toDate8], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate8) && DateTime.TryParseExact(config[configs.fromDate9], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate9) && DateTime.TryParseExact(config[configs.toDate9], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate9)
                                                && DateTime.TryParseExact(config[configs.fromDate10], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate10) && DateTime.TryParseExact(config[configs.toDate10], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate10)
                                                && DateTime.TryParseExact(config[configs.fromDate11], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate11) && DateTime.TryParseExact(config[configs.toDate11], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate11)
                                                && DateTime.TryParseExact(config[configs.fromDate12], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate12) && DateTime.TryParseExact(config[configs.toDate12], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate12)
                                                && DateTime.TryParseExact(config[configs.fromDate13], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate13) && DateTime.TryParseExact(config[configs.toDate13], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate13)
                                                && DateTime.TryParseExact(config[configs.fromDate14], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate14) && DateTime.TryParseExact(config[configs.toDate14], fommatDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate14))
                                        {
                                            for (DateTime date = startDate7; date <= endDate1; date = date.AddDays(1))
                                            {

                                                if (group6Date >= startDate7 && group6Date <= endDate1)
                                                {
                                                    if (group7Value.Contains(configs.Session1) || group7Value.Contains(configs.Session2))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");
                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;
                                                                    string sheetName = config[configs.sheetName];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        int beginrow1 = 0;
                                                                        int beginrow4 = 0;
                                                                        int beginrow2 = 0;
                                                                        if (group6Date >= startDate1 && group6Date <= endDate1)
                                                                        {
                                                                            beginrow1 = Int32.Parse(config[configs.beginrow1]); //8
                                                                            beginrow2 = Int32.Parse(config[configs.beginrow1]); //8
                                                                            beginrow4 = Int32.Parse(config[configs.beginrow1]); //8
                                                                        }
                                                                        else if (group6Date >= startDate2 && group6Date <= endDate2)
                                                                        {
                                                                            beginrow1 = Int32.Parse(config[configs.beginrow2]); //15
                                                                            beginrow2 = Int32.Parse(config[configs.beginrow1]); //8
                                                                            beginrow4 = Int32.Parse(config[configs.beginrow2]);  //15
                                                                        }
                                                                        else if (group6Date >= startDate3 && group6Date <= endDate3)
                                                                        {
                                                                            beginrow1 = Int32.Parse(config[configs.beginrow2]); //15
                                                                            beginrow2 = Int32.Parse(config[configs.beginrow2]); //15
                                                                            beginrow4 = Int32.Parse(config[configs.beginrow1]);  //8
                                                                        }
                                                                        else if ((group6Date >= startDate4 && group6Date <= endDate4) || (group6Date >= startDate6 && group6Date <= endDate6))
                                                                        {
                                                                            beginrow1 = Int32.Parse(config[configs.beginrow3]);  //11
                                                                            beginrow2 = Int32.Parse(config[configs.beginrow3]); //11
                                                                            beginrow4 = Int32.Parse(config[configs.beginrow1]);  //8
                                                                        }

                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {
                                                                            try
                                                                            {
                                                                                if (sheet == "1")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen1(dataSet);
                                                                                    string ColumOfDB = config[configs.data1col1];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table1 = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        table1.Columns.Add(column1);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["1"];

                                                                                    string beginRow1 = config[configs.databeginRow1];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = beginrow1; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.PriorDayClose = view;
                                                                                            }
                                                                                            else tbl.PriorDayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionClose = view;
                                                                                            }
                                                                                            else tbl.SessionClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Change = view;
                                                                                            }
                                                                                            else tbl.Change = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table1.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.PriorDayClose, tbl.SessionClose, tbl.Change, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table1, config[configs.tableSession1]);
                                                                                }
                                                                                else if (sheet == "1.CW")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen1CW(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col2];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table1cw = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table1cw.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["1.CW"];
                                                                                    string beginRow1 = config[configs.databeginRow2];

                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow1; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;
                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.PriorDayClose = view;
                                                                                            }
                                                                                            else tbl.PriorDayClose = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionClose = view;
                                                                                            }
                                                                                            else tbl.SessionClose = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Change = view;
                                                                                            }
                                                                                            else tbl.Change = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table1cw.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.PriorDayClose, tbl.SessionClose, tbl.Change, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table1cw, config[configs.tableSession1CW]);
                                                                                }
                                                                                else if (sheet == "2")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen2(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col3];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table2 = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table2.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["2"];
                                                                                    string beginRow1 = config[configs.databeginRow3];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow2; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.PriorDayClose = view;
                                                                                            }
                                                                                            else tbl.PriorDayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionHigh = view;
                                                                                            }
                                                                                            else tbl.SessionHigh = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionAverage = view;
                                                                                            }
                                                                                            else tbl.SessionAverage = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionLow = view;
                                                                                            }
                                                                                            else tbl.SessionLow = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionClose = view;
                                                                                            }
                                                                                            else tbl.SessionClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table2.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.PriorDayClose, tbl.SessionHigh, tbl.SessionAverage, tbl.SessionLow, tbl.SessionClose, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table2, config[configs.tableSession2]);
                                                                                }
                                                                                else if (sheet == "2.CW")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen2CW(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col4];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table2cw = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table2cw.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["2.CW"];
                                                                                    string beginRow1 = config[configs.databeginRow4];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow1; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();


                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.PriorDayClose = view;
                                                                                            }
                                                                                            else tbl.PriorDayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionHigh = view;
                                                                                            }
                                                                                            else tbl.SessionHigh = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionAverage = view;
                                                                                            }
                                                                                            else tbl.SessionAverage = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionLow = view;
                                                                                            }
                                                                                            else tbl.SessionLow = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.SessionClose = view;
                                                                                            }
                                                                                            else tbl.SessionClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table2cw.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.PriorDayClose, tbl.SessionHigh, tbl.SessionAverage, tbl.SessionLow, tbl.SessionClose, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table2cw, config[configs.tableSession2CW]);
                                                                                }
                                                                                else if (sheet == "2.ODD_OM")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen2ODD_OM(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col5];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table2oddom = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table2oddom.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["2.ODD_OM"];
                                                                                    string beginRow1 = config[configs.databeginRow5];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow1; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table2oddom.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }

                                                                                    insert.InsertDB(table2oddom, config[configs.tableSession2ODD]);


                                                                                }
                                                                                else if (sheet == "3")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen3(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col6];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table3 = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table3.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["3"];
                                                                                    string beginRow1 = config[configs.databeginRow6];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow1; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.PriorDayClose = view;
                                                                                            }
                                                                                            else tbl.PriorDayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TodayClose = view;
                                                                                            }
                                                                                            else tbl.TodayClose = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Change = view;
                                                                                            }
                                                                                            else tbl.Change = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Totalvolume = view;
                                                                                            }
                                                                                            else tbl.Totalvolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Totalvalue = view;
                                                                                            }
                                                                                            else tbl.Totalvalue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[8]].ToString(), out view))
                                                                                            {
                                                                                                tbl.ListedShares = view;
                                                                                            }
                                                                                            else tbl.ListedShares = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[9]].ToString(), out view))
                                                                                            {
                                                                                                tbl.OutstandingShares = view;
                                                                                            }
                                                                                            else tbl.OutstandingShares = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[10]].ToString(), out view))
                                                                                            {
                                                                                                tbl.AdjustedOutstandingShares = view;
                                                                                            }
                                                                                            else tbl.AdjustedOutstandingShares = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[11]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Marketcap = view;
                                                                                            }
                                                                                            else tbl.Marketcap = 0;

                                                                                            table3.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.PriorDayClose, tbl.TodayClose, tbl.Change, tbl.TradingVolume, tbl.TradingValue, tbl.Totalvolume, tbl.Totalvalue, tbl.ListedShares, tbl.OutstandingShares, tbl.AdjustedOutstandingShares, tbl.Marketcap);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table3, config[configs.tableSession3]);

                                                                                }
                                                                                else if (sheet == "3.CW")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen3CW(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col7];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table3cw = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table3cw.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["3.CW"];
                                                                                    string beginRow1 = config[configs.databeginRow7];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow1; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.PriorDayClose = view;
                                                                                            }
                                                                                            else tbl.PriorDayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TodayClose = view;
                                                                                            }
                                                                                            else tbl.TodayClose = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Change = view;
                                                                                            }
                                                                                            else tbl.Change = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Totalvolume = view;
                                                                                            }
                                                                                            else tbl.Totalvolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Totalvalue = view;
                                                                                            }
                                                                                            else tbl.Totalvalue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[8]].ToString(), out view))
                                                                                            {
                                                                                                tbl.ListedShares = view;
                                                                                            }
                                                                                            else tbl.ListedShares = 0;

                                                                                            table3cw.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.PriorDayClose, tbl.TodayClose, tbl.Change, tbl.TradingVolume, tbl.TradingValue, tbl.Totalvolume, tbl.Totalvalue, tbl.ListedShares);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table3cw, config[configs.tableSession3CW]);


                                                                                }
                                                                                else if (sheet == "4")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen4(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col8];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table4 = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table4.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["4"];
                                                                                    string beginRow1 = config[configs.databeginRow8];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = beginrow4; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table4.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table4, config[configs.tableSession4]);

                                                                                }
                                                                                else if (sheet == "4.CW")
                                                                                {
                                                                                    var dataSetX = configTable.DatTen4CW(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col9];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table4cw = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table4cw.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["4.CW"];
                                                                                    string beginRow1 = config[configs.databeginRow9];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = Int32.Parse(config[configs.beginrow1]); y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();


                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table4cw.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table4cw, config[configs.tableSession4CW]);

                                                                                }
                                                                                else if (sheet == "4.ODD_PT")
                                                                                {
                                                                                    var dataSetX = configTable.DatTenODD_PT(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col10];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable table4oddpt = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        table4oddpt.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["4.ODD_PT"];
                                                                                    string beginRow1 = config[configs.databeginRow10];
                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = Int32.Parse(config[configs.beginrow1]); y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                                        {
                                                                                            tbl.StockCode = (dataTable.Rows[y][column1[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;

                                                                                            table4oddpt.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.StockCode, tbl.TradingVolume, tbl.TradingValue);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(table4oddpt, config[configs.tableSession4ODD]);
                                                                                }
                                                                                else if (sheet == "HOSEINDEX")
                                                                                {
                                                                                    var dataSetX = configTable.DatTenHOSEINDEX(dataSet);

                                                                                    string ColumOfDB = config[configs.data1col11];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tableHOSEINDEX = new DataTable();
                                                                                    foreach (string column in columns)
                                                                                    {
                                                                                        tableHOSEINDEX.Columns.Add(column);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["HOSEINDEX"];
                                                                                    string beginRow1 = config[configs.databeginCell11];

                                                                                    string[] column1 = beginRow1.Split(',');

                                                                                    for (int y = Int32.Parse(config[configs.databeginRow11]); y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETable tbl = new ETable();
                                                                                        tbl.CreateDate = DateTime.Now;
                                                                                        tbl.TransDate = group6Date;
                                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                        {
                                                                                            tbl.IndexName = (string)dataTable.Rows[y][column1[0]];

                                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                            {
                                                                                                tbl.OpenIndexValue = view;
                                                                                            }
                                                                                            else tbl.OpenIndexValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                            {
                                                                                                tbl.CloseIndexValue = view;
                                                                                            }
                                                                                            else tbl.CloseIndexValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                            {
                                                                                                tbl.High = view;
                                                                                            }
                                                                                            else tbl.High = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Low = view;
                                                                                            }
                                                                                            else tbl.Low = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                            {
                                                                                                tbl.UpDown = view;
                                                                                            }
                                                                                            else tbl.UpDown = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Change = view;
                                                                                            }
                                                                                            else tbl.Change = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingVolume = view;
                                                                                            }
                                                                                            else tbl.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[8]].ToString(), out view))
                                                                                            {
                                                                                                tbl.TradingValue = view;
                                                                                            }
                                                                                            else tbl.TradingValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column1[9]].ToString(), out view))
                                                                                            {
                                                                                                tbl.Marketcap = view;
                                                                                            }
                                                                                            else tbl.Marketcap = 0;

                                                                                            tableHOSEINDEX.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.IndexName, tbl.OpenIndexValue, tbl.CloseIndexValue, tbl.High, tbl.Low, tbl.UpDown, tbl.Change, tbl.TradingVolume, tbl.TradingValue, tbl.Marketcap);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tableHOSEINDEX, config[configs.tableSessionHOSE]);
                                                                                }
                                                                            }
                                                                            catch (Exception e)
                                                                            {
                                                                                Console.WriteLine($"Don't have sheet {sheet} : error {e.Message}");
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.Basic1) || group7Value.Contains(configs.Basic2))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;
                                                                    var dataSetX = configTable.DatTenBasic1(dataSet);
                                                                    string ColumOfDB = config[configs.data2col1];

                                                                    string[] columns = ColumOfDB.Split(',');

                                                                    DataTable table1basic = new DataTable();
                                                                    foreach (string column1 in columns)
                                                                    {
                                                                        table1basic.Columns.Add(column1);
                                                                    }
                                                                    DataTable dataTable = dataSetX.Tables["1"];

                                                                    string beginRow1 = config[configs.data2beginRow1];

                                                                    string[] column = beginRow1.Split(',');

                                                                    for (int y = Int32.Parse(config[configs.set2beginrow1]); y < dataTable.Rows.Count; y++)
                                                                    {
                                                                        ETableBasic tblbasic = new ETableBasic();
                                                                        tblbasic.CreateDate = DateTime.Now;
                                                                        tblbasic.TransDate = group6Date;

                                                                        if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                        {
                                                                            tblbasic.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                            {
                                                                                tblbasic.PriorDayClose = view;
                                                                            }
                                                                            else tblbasic.PriorDayClose = 0;

                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                            {
                                                                                tblbasic.wRecordHigh = view;
                                                                            }
                                                                            else tblbasic.wRecordHigh = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                            {
                                                                                tblbasic.wRecordLow = view;
                                                                            }
                                                                            else tblbasic.wRecordLow = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                            {
                                                                                tblbasic.AverageOutstandingShares = view;
                                                                            }
                                                                            else tblbasic.AverageOutstandingShares = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                            {
                                                                                tblbasic.PrimaryEPS = view;
                                                                            }
                                                                            else tblbasic.PrimaryEPS = 0;
                                                                            if (dataTable.Rows[y][column[6]].ToString() != "")
                                                                            {
                                                                                tblbasic.Notes = (dataTable.Rows[y][column[6]]).ToString();
                                                                            }
                                                                            else tblbasic.Notes = "NULL";
                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                            {
                                                                                tblbasic.AdjustedEPS = view;
                                                                            }
                                                                            else tblbasic.AdjustedEPS = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                            {
                                                                                tblbasic.PE = view;
                                                                            }
                                                                            else tblbasic.PE = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                            {
                                                                                tblbasic.Dividend = view;
                                                                            }
                                                                            else tblbasic.Dividend = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                            {
                                                                                tblbasic.DividendMarketPrice = view;
                                                                            }
                                                                            else tblbasic.DividendMarketPrice = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                            {
                                                                                tblbasic.ReturnOnTotalAssets = view;
                                                                            }
                                                                            else tblbasic.ReturnOnTotalAssets = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                            {
                                                                                tblbasic.ReturnOnEquity = view;
                                                                            }
                                                                            else tblbasic.ReturnOnEquity = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                            {
                                                                                tblbasic.ListedShares = view;
                                                                            }
                                                                            else tblbasic.ListedShares = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                            {
                                                                                tblbasic.OutstandingShares = view;
                                                                            }
                                                                            else tblbasic.OutstandingShares = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                            {
                                                                                tblbasic.ChangeOutstandingShares = view;
                                                                            }
                                                                            else tblbasic.ChangeOutstandingShares = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                            {
                                                                                tblbasic.AdjustedOutstandingShares = view;
                                                                            }
                                                                            else tblbasic.AdjustedOutstandingShares = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[17]].ToString(), out view))
                                                                            {
                                                                                tblbasic.TurnoverRatio = view;
                                                                            }
                                                                            else tblbasic.TurnoverRatio = 0;

                                                                            table1basic.Rows.Add(tblbasic.TransDate, tblbasic.CreateDate, tblbasic.StockCode, tblbasic.PriorDayClose, tblbasic.wRecordHigh, tblbasic.wRecordLow, tblbasic.AverageOutstandingShares, tblbasic.PrimaryEPS, tblbasic.Notes, tblbasic.AdjustedEPS, tblbasic.PE, tblbasic.Dividend, tblbasic.DividendMarketPrice, tblbasic.ReturnOnTotalAssets, tblbasic.ReturnOnEquity, tblbasic.ListedShares, tblbasic.OutstandingShares, tblbasic.ChangeOutstandingShares, tblbasic.AdjustedOutstandingShares, tblbasic.TurnoverRatio);
                                                                        }
                                                                    }
                                                                    insert.InsertDB(table1basic, config[configs.tableBasic]);
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.MarketCap))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;
                                                                    var dataSetX = configTable.DatTenMatketCap(dataSet);
                                                                    string ColumOfDB = config[configs.marketcol];

                                                                    string[] columns = ColumOfDB.Split(',');

                                                                    DataTable table1matket = new DataTable();
                                                                    foreach (string column1 in columns)
                                                                    {
                                                                        table1matket.Columns.Add(column1);
                                                                    }
                                                                    DataTable dataTable = dataSetX.Tables["1"];

                                                                    string beginRow1 = config[configs.marketcell1];

                                                                    string[] column = beginRow1.Split(',');

                                                                    for (int y = Int32.Parse(config[configs.marketrow1]); y < dataTable.Rows.Count; y++)
                                                                    {
                                                                        ETableMatketCap tblMatket = new ETableMatketCap();
                                                                        tblMatket.CreateDate = DateTime.Now;
                                                                        tblMatket.TransDate = group6Date;

                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                        {
                                                                            tblMatket.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                            {
                                                                                tblMatket.Session_Year_Count = view;
                                                                            }
                                                                            else tblMatket.Session_Year_Count = 0;

                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                            {
                                                                                tblMatket.TradingVolume_Year = view;
                                                                            }
                                                                            else tblMatket.TradingVolume_Year = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                            {
                                                                                tblMatket.AvgSession1 = view;
                                                                            }
                                                                            else tblMatket.AvgSession1 = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                            {
                                                                                tblMatket.TradingValue_Year = view;
                                                                            }
                                                                            else tblMatket.TradingValue_Year = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                            {
                                                                                tblMatket.AvgSession2 = view;
                                                                            }
                                                                            else tblMatket.AvgSession2 = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                            {
                                                                                tblMatket.Price = view;
                                                                            }
                                                                            else tblMatket.Price = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                            {
                                                                                tblMatket.KLNY = view;
                                                                            }
                                                                            else tblMatket.KLNY = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                            {
                                                                                tblMatket.KLNY_Current = view;
                                                                            }
                                                                            else tblMatket.KLNY_Current = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                            {
                                                                                tblMatket.GTVH = view;
                                                                            }
                                                                            else tblMatket.GTVH = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                            {
                                                                                tblMatket.SpeedChange = view;
                                                                            }
                                                                            else tblMatket.SpeedChange = 0;

                                                                            table1matket.Rows.Add(tblMatket.TransDate, tblMatket.CreateDate, tblMatket.StockCode, tblMatket.Session_Year_Count, tblMatket.TradingVolume_Year, tblMatket.AvgSession1, tblMatket.TradingValue_Year, tblMatket.AvgSession2, tblMatket.Price, tblMatket.KLNY, tblMatket.KLNY_Current, tblMatket.GTVH, tblMatket.SpeedChange);
                                                                        }
                                                                        else if (dataTable.Rows[y][column[0]].ToString() == "")
                                                                        {
                                                                            break;
                                                                        }
                                                                    }
                                                                    insert.InsertDB(table1matket, config[configs.tableMarket]);
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.CC))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;
                                                                    string sheetName = config[configs.ordermatchingSheetname];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {

                                                                            if (sheet == "2")
                                                                            {
                                                                                var dataSetX = configTable.DatTenOrrderMatching(dataSet);
                                                                                string ColumOfDB = config[configs.ordermatchingcol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table2ordermatch = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    table2ordermatch.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["2"];

                                                                                string beginRow1 = config[configs.ordermatchingcell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                // check -6
                                                                                for (int y = Int32.Parse(config[configs.ordermatchingrow]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableCC tblOrderMatching = new ETableCC();
                                                                                    tblOrderMatching.CreateDate = DateTime.Now;
                                                                                    tblOrderMatching.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length <= 5)
                                                                                    {
                                                                                        tblOrderMatching.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.BuyingOrders = view;
                                                                                        }
                                                                                        else tblOrderMatching.BuyingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.BuyingVolume = view;
                                                                                        }
                                                                                        else tblOrderMatching.BuyingVolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.SellingOrders = view;
                                                                                        }
                                                                                        else tblOrderMatching.SellingOrders = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.SellingVolume = view;
                                                                                        }
                                                                                        else tblOrderMatching.SellingVolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.Change = view;
                                                                                        }
                                                                                        else tblOrderMatching.Change = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.CeilingPrice = view;
                                                                                        }
                                                                                        else tblOrderMatching.CeilingPrice = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.FloorPrice = view;
                                                                                        }
                                                                                        else tblOrderMatching.FloorPrice = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.BestBidPrice = view;
                                                                                        }
                                                                                        else tblOrderMatching.BestBidPrice = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.BestOfferPrice = view;
                                                                                        }
                                                                                        else tblOrderMatching.BestOfferPrice = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                        {
                                                                                            tblOrderMatching.Spreads = view;
                                                                                        }
                                                                                        else tblOrderMatching.Spreads = 0;

                                                                                        table2ordermatch.Rows.Add(tblOrderMatching.TransDate, tblOrderMatching.CreateDate, tblOrderMatching.StockCode, tblOrderMatching.BuyingOrders, tblOrderMatching.BuyingVolume, tblOrderMatching.SellingOrders, tblOrderMatching.SellingVolume, tblOrderMatching.Change, tblOrderMatching.CeilingPrice, tblOrderMatching.FloorPrice, tblOrderMatching.BestBidPrice, tblOrderMatching.BestOfferPrice, tblOrderMatching.Spreads);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table2ordermatch, config[configs.tableOrder]);
                                                                            }

                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.CP) || group7Value.Contains(configs.TCNY))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                if (group6Date >= startDate8 && group6Date <= endDate8)
                                                                {
                                                                    var dataSet = reader.AsDataSet();
                                                                    if (dataSet != null)
                                                                    {
                                                                        ConfigTable configTable = new ConfigTable();

                                                                        float view;
                                                                        var dataSetX = configTable.DatTenBasicOld2(dataSet);
                                                                        string ColumOfDB = config[configs.data2old2col];

                                                                        string[] columns = ColumOfDB.Split(',');

                                                                        DataTable table2basic = new DataTable();
                                                                        foreach (string column1 in columns)
                                                                        {
                                                                            table2basic.Columns.Add(column1);
                                                                        }
                                                                        DataTable dataTable = dataSetX.Tables["1"];

                                                                        string beginRow1 = config[configs.dataold2cell];

                                                                        string[] column = beginRow1.Split(',');

                                                                        for (int y = Int32.Parse(config[configs.dataold2row]); y < dataTable.Rows.Count; y++)
                                                                        {
                                                                            ETableBasic tblbasic = new ETableBasic();
                                                                            tblbasic.CreateDate = DateTime.Now;
                                                                            tblbasic.TransDate = group6Date;

                                                                            if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString().Length < 15 && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                            {
                                                                                tblbasic.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.wRecordHigh = view;
                                                                                }
                                                                                else tblbasic.wRecordHigh = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.wRecordLow = view;
                                                                                }
                                                                                else tblbasic.wRecordLow = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AverageOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.AverageOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PrimaryEPS = view;
                                                                                }
                                                                                else tblbasic.PrimaryEPS = 0;
                                                                                if (dataTable.Rows[y][column[5]].ToString() != "")
                                                                                {
                                                                                    tblbasic.Notes = (dataTable.Rows[y][column[5]]).ToString();
                                                                                }
                                                                                else tblbasic.Notes = "NULL";
                                                                                if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AdjustedEPS = view;
                                                                                }
                                                                                else tblbasic.AdjustedEPS = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PE = view;
                                                                                }
                                                                                else tblbasic.PE = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.Dividend = view;
                                                                                }
                                                                                else tblbasic.Dividend = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.DividendMarketPrice = view;
                                                                                }
                                                                                else tblbasic.DividendMarketPrice = 0;

                                                                                if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.ListedShares = view;
                                                                                }
                                                                                else tblbasic.ListedShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.OutstandingShares = view;
                                                                                }
                                                                                else tblbasic.OutstandingShares = 0;

                                                                                if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AdjustedOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.AdjustedOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.MtkCap = view;
                                                                                }
                                                                                else tblbasic.MtkCap = 0;

                                                                                table2basic.Rows.Add(tblbasic.TransDate, tblbasic.CreateDate, tblbasic.StockCode, tblbasic.wRecordHigh, tblbasic.wRecordLow, tblbasic.AverageOutstandingShares, tblbasic.PrimaryEPS, tblbasic.Notes, tblbasic.AdjustedEPS, tblbasic.PE, tblbasic.Dividend, tblbasic.DividendMarketPrice, tblbasic.ListedShares, tblbasic.OutstandingShares, tblbasic.AdjustedOutstandingShares, tblbasic.MtkCap);
                                                                            }
                                                                        }
                                                                        insert.InsertDB(table2basic, config[configs.tableBasic1]);

                                                                    }
                                                                    else
                                                                    {
                                                                        throw new Exception(configs.ERROR);
                                                                    }

                                                                }
                                                                else if (group6Date >= startDate10 && group6Date <= endDate10)
                                                                {
                                                                    var dataSet = reader.AsDataSet();
                                                                    if (dataSet != null)
                                                                    {
                                                                        ConfigTable configTable = new ConfigTable();

                                                                        float view;
                                                                        var dataSetX = configTable.DatTenBasicOld3(dataSet);
                                                                        string ColumOfDB = config[configs.data2old3col];

                                                                        string[] columns = ColumOfDB.Split(',');

                                                                        DataTable table2basic = new DataTable();
                                                                        foreach (string column1 in columns)
                                                                        {
                                                                            table2basic.Columns.Add(column1);
                                                                        }
                                                                        DataTable dataTable = dataSetX.Tables["1"];

                                                                        string beginRow1 = config[configs.dataold3cell];

                                                                        string[] column = beginRow1.Split(',');
                                                                        for (int y = Int32.Parse(config[configs.dataold3row]); y < dataTable.Rows.Count; y++)
                                                                        {
                                                                            ETableBasic tblbasic = new ETableBasic();
                                                                            tblbasic.CreateDate = DateTime.Now;
                                                                            tblbasic.TransDate = group6Date;

                                                                            if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                            {
                                                                                tblbasic.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.wRecordHigh = view;
                                                                                }
                                                                                else tblbasic.wRecordHigh = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.wRecordLow = view;
                                                                                }
                                                                                else tblbasic.wRecordLow = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AverageOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.AverageOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PrimaryEPS = view;
                                                                                }
                                                                                else tblbasic.PrimaryEPS = 0;
                                                                                if (dataTable.Rows[y][column[5]].ToString() != "")
                                                                                {
                                                                                    tblbasic.Notes = (dataTable.Rows[y][column[5]]).ToString();
                                                                                }
                                                                                else tblbasic.Notes = "NULL";
                                                                                if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AdjustedEPS = view;
                                                                                }
                                                                                else tblbasic.AdjustedEPS = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PE = view;
                                                                                }
                                                                                else tblbasic.PE = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.Dividend = view;
                                                                                }
                                                                                else tblbasic.Dividend = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.DividendMarketPrice = view;
                                                                                }
                                                                                else tblbasic.DividendMarketPrice = 0;

                                                                                if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.ListedShares = view;
                                                                                }
                                                                                else tblbasic.ListedShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.OutstandingShares = view;
                                                                                }
                                                                                else tblbasic.OutstandingShares = 0;


                                                                                if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.MtkCap = view;
                                                                                }
                                                                                else tblbasic.MtkCap = 0;

                                                                                table2basic.Rows.Add(tblbasic.TransDate, tblbasic.CreateDate, tblbasic.StockCode, tblbasic.wRecordHigh, tblbasic.wRecordLow, tblbasic.AverageOutstandingShares, tblbasic.PrimaryEPS, tblbasic.Notes, tblbasic.AdjustedEPS, tblbasic.PE, tblbasic.Dividend, tblbasic.DividendMarketPrice, tblbasic.ListedShares, tblbasic.OutstandingShares, tblbasic.MtkCap);
                                                                            }
                                                                        }
                                                                        StringBuilder sb = new StringBuilder();
                                                                        sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableBasic1]).Append(" (").Append(config[configs.data2old3col]).Append(") ").Append(configs.valueTbl).Append(" ");
                                                                        foreach (DataRow row in table2basic.Rows)
                                                                        {
                                                                            sb.Append("(");
                                                                            sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                            sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                            sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                            sb.Append(row["wRecordHigh"]).Append(",");
                                                                            sb.Append(row["wRecordLow"]).Append(",");
                                                                            sb.Append(row["AverageOutstandingShares"]).Append(",");
                                                                            sb.Append(row["PrimaryEPS"]).Append(",");
                                                                            sb.Append("'" + row["Notes"] + "'").Append(",");
                                                                            sb.Append(row["AdjustedEPS"]).Append(",");
                                                                            sb.Append(row["PE"]).Append(",");
                                                                            sb.Append(row["Dividend"]).Append(",");
                                                                            sb.Append(row["DividendMarketPrice"]).Append(",");
                                                                            sb.Append(row["ListedShares"]).Append(",");
                                                                            sb.Append(row["OutstandingShares"]).Append(",");
                                                                            sb.Append(row["MtkCap"]);
                                                                            sb.Append("),");
                                                                        }
                                                                        command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                        command.ExecuteNonQuery();
                                                                    }
                                                                    else
                                                                    {
                                                                        throw new Exception(configs.ERROR);
                                                                    }
                                                                }
                                                                else if ((group6Date >= startDate9 && group6Date <= endDate9) || (group6Date >= startDate11 && group6Date <= endDate11))
                                                                {
                                                                    var dataSet = reader.AsDataSet();
                                                                    if (dataSet != null)
                                                                    {
                                                                        ConfigTable configTable = new ConfigTable();

                                                                        float view;
                                                                        var dataSetX = configTable.DatTenBasicOld1(dataSet);
                                                                        string ColumOfDB = config[configs.data2old1col];

                                                                        string[] columns = ColumOfDB.Split(',');

                                                                        DataTable table2basic = new DataTable();
                                                                        foreach (string column1 in columns)
                                                                        {
                                                                            table2basic.Columns.Add(column1);
                                                                        }
                                                                        DataTable dataTable = dataSetX.Tables["1"];

                                                                        string beginRow1 = config[configs.dataold1cell];

                                                                        string[] column = beginRow1.Split(',');

                                                                        for (int y = Int32.Parse(config[configs.dataold1row]); y < dataTable.Rows.Count; y++)
                                                                        {
                                                                            ETableBasic tblbasic = new ETableBasic();
                                                                            tblbasic.CreateDate = DateTime.Now;
                                                                            tblbasic.TransDate = group6Date;

                                                                            if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                            {
                                                                                tblbasic.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AverageOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.AverageOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PrimaryEPS = view;
                                                                                }
                                                                                else tblbasic.PrimaryEPS = 0;
                                                                                if (dataTable.Rows[y][column[3]].ToString() != "")
                                                                                {
                                                                                    tblbasic.Notes = (dataTable.Rows[y][column[3]]).ToString();
                                                                                }
                                                                                else tblbasic.Notes = "NULL";
                                                                                if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AdjustedEPS = view;
                                                                                }
                                                                                else tblbasic.AdjustedEPS = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PE = view;
                                                                                }
                                                                                else tblbasic.PE = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.Dividend = view;
                                                                                }
                                                                                else tblbasic.Dividend = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.DividendMarketPrice = view;
                                                                                }
                                                                                else tblbasic.DividendMarketPrice = 0;

                                                                                if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.ListedShares = view;
                                                                                }
                                                                                else tblbasic.ListedShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.OutstandingShares = view;
                                                                                }
                                                                                else tblbasic.OutstandingShares = 0;


                                                                                if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.MtkCap = view;
                                                                                }
                                                                                else tblbasic.MtkCap = 0;

                                                                                table2basic.Rows.Add(tblbasic.TransDate, tblbasic.CreateDate, tblbasic.StockCode, tblbasic.AverageOutstandingShares, tblbasic.PrimaryEPS, tblbasic.Notes, tblbasic.AdjustedEPS, tblbasic.PE, tblbasic.Dividend, tblbasic.DividendMarketPrice, tblbasic.ListedShares, tblbasic.OutstandingShares, tblbasic.MtkCap);
                                                                            }
                                                                        }
                                                                        StringBuilder sb = new StringBuilder();
                                                                        sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableBasic1]).Append(" (").Append(config[configs.data2old1col]).Append(") ").Append(configs.valueTbl).Append(" ");
                                                                        foreach (DataRow row in table2basic.Rows)
                                                                        {
                                                                            sb.Append("(");
                                                                            sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                            sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                            sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                            sb.Append(row["AverageOutstandingShares"]).Append(",");
                                                                            sb.Append(row["PrimaryEPS"]).Append(",");
                                                                            sb.Append("'" + row["Notes"] + "'").Append(",");
                                                                            sb.Append(row["AdjustedEPS"]).Append(",");
                                                                            sb.Append(row["PE"]).Append(",");
                                                                            sb.Append(row["Dividend"]).Append(",");
                                                                            sb.Append(row["DividendMarketPrice"]).Append(",");
                                                                            sb.Append(row["ListedShares"]).Append(",");
                                                                            sb.Append(row["OutstandingShares"]).Append(",");
                                                                            sb.Append(row["MtkCap"]);
                                                                            sb.Append("),");
                                                                        }
                                                                        command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                        command.ExecuteNonQuery();
                                                                    }
                                                                    else
                                                                    {
                                                                        throw new Exception(configs.ERROR);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    var dataSet = reader.AsDataSet();
                                                                    if (dataSet != null)
                                                                    {
                                                                        ConfigTable configTable = new ConfigTable();

                                                                        float view;
                                                                        var dataSetX = configTable.DatTenBasicOld(dataSet);
                                                                        string ColumOfDB = config[configs.data2oldcol];

                                                                        string[] columns = ColumOfDB.Split(',');

                                                                        DataTable table2basic = new DataTable();
                                                                        foreach (string column1 in columns)
                                                                        {
                                                                            table2basic.Columns.Add(column1);
                                                                        }
                                                                        DataTable dataTable = dataSetX.Tables["1"];

                                                                        string beginRow1 = config[configs.dataoldcell];

                                                                        string[] column = beginRow1.Split(',');

                                                                        for (int y = Int32.Parse(config[configs.dataoldrow]); y < dataTable.Rows.Count; y++)
                                                                        {
                                                                            ETableBasic tblbasic = new ETableBasic();
                                                                            tblbasic.CreateDate = DateTime.Now;
                                                                            tblbasic.TransDate = group6Date;

                                                                            if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                            {
                                                                                tblbasic.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.wRecordHigh = view;
                                                                                }
                                                                                else tblbasic.wRecordHigh = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.wRecordLow = view;
                                                                                }
                                                                                else tblbasic.wRecordLow = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AverageOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.AverageOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PrimaryEPS = view;
                                                                                }
                                                                                else tblbasic.PrimaryEPS = 0;
                                                                                if (dataTable.Rows[y][column[5]].ToString() != "")
                                                                                {
                                                                                    tblbasic.Notes = (dataTable.Rows[y][column[5]]).ToString();
                                                                                }
                                                                                else tblbasic.Notes = "NULL";
                                                                                if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AdjustedEPS = view;
                                                                                }
                                                                                else tblbasic.AdjustedEPS = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.PE = view;
                                                                                }
                                                                                else tblbasic.PE = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.Dividend = view;
                                                                                }
                                                                                else tblbasic.Dividend = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.DividendMarketPrice = view;
                                                                                }
                                                                                else tblbasic.DividendMarketPrice = 0;

                                                                                if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.ListedShares = view;
                                                                                }
                                                                                else tblbasic.ListedShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.OutstandingShares = view;
                                                                                }
                                                                                else tblbasic.OutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.ChangeOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.ChangeOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.AdjustedOutstandingShares = view;
                                                                                }
                                                                                else tblbasic.AdjustedOutstandingShares = 0;
                                                                                if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                {
                                                                                    tblbasic.MtkCap = view;
                                                                                }
                                                                                else tblbasic.MtkCap = 0;

                                                                                table2basic.Rows.Add(tblbasic.TransDate, tblbasic.CreateDate, tblbasic.StockCode, tblbasic.wRecordHigh, tblbasic.wRecordLow, tblbasic.AverageOutstandingShares, tblbasic.PrimaryEPS, tblbasic.Notes, tblbasic.AdjustedEPS, tblbasic.PE, tblbasic.Dividend, tblbasic.DividendMarketPrice, tblbasic.ListedShares, tblbasic.OutstandingShares, tblbasic.ChangeOutstandingShares, tblbasic.AdjustedOutstandingShares, tblbasic.MtkCap);
                                                                            }
                                                                        }
                                                                        insert.InsertDB(table2basic, config[configs.tableBasic1]);
                                                                    }
                                                                    else
                                                                    {
                                                                        throw new Exception(configs.ERROR);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.Indices))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;

                                                                    var dataSetX = configTable.DatTenHOSEINDEX(dataSet);

                                                                    string ColumOfDB = config[configs.data1col11];

                                                                    string[] columns = ColumOfDB.Split(',');

                                                                    DataTable tableHOSEINDEX = new DataTable();
                                                                    foreach (string column in columns)
                                                                    {
                                                                        tableHOSEINDEX.Columns.Add(column);
                                                                    }
                                                                    DataTable dataTable = dataSetX.Tables["HOSEINDEX"];
                                                                    string beginRow1 = config[configs.databeginCell11];

                                                                    string[] column1 = beginRow1.Split(',');

                                                                    for (int y = Int32.Parse(config[configs.databeginRow11]); y < dataTable.Rows.Count; y++)
                                                                    {
                                                                        ETable tbl = new ETable();
                                                                        tbl.CreateDate = DateTime.Now;
                                                                        tbl.TransDate = group6Date;
                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                        {
                                                                            tbl.IndexName = (string)dataTable.Rows[y][column1[0]];

                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                            {
                                                                                tbl.OpenIndexValue = view;
                                                                            }
                                                                            else tbl.OpenIndexValue = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                            {
                                                                                tbl.CloseIndexValue = view;
                                                                            }
                                                                            else tbl.CloseIndexValue = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                            {
                                                                                tbl.High = view;
                                                                            }
                                                                            else tbl.High = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                            {
                                                                                tbl.Low = view;
                                                                            }
                                                                            else tbl.Low = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                            {
                                                                                tbl.UpDown = view;
                                                                            }
                                                                            else tbl.UpDown = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                            {
                                                                                tbl.Change = view;
                                                                            }
                                                                            else tbl.Change = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                            {
                                                                                tbl.TradingVolume = view;
                                                                            }
                                                                            else tbl.TradingVolume = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[8]].ToString(), out view))
                                                                            {
                                                                                tbl.TradingValue = view;
                                                                            }
                                                                            else tbl.TradingValue = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[9]].ToString(), out view))
                                                                            {
                                                                                tbl.Marketcap = view;
                                                                            }
                                                                            else tbl.Marketcap = 0;


                                                                            tableHOSEINDEX.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.IndexName, tbl.OpenIndexValue, tbl.CloseIndexValue, tbl.High, tbl.Low, tbl.UpDown, tbl.Change, tbl.TradingVolume, tbl.TradingValue, tbl.Marketcap);
                                                                        }
                                                                    }
                                                                    insert.InsertDB(tableHOSEINDEX, config[configs.tableSessionHOSE]);

                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.Corporate))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");
                                                                var dataSet = reader.AsDataSet();
                                                                // dặt ten 
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;
                                                                    DateTime exDate;
                                                                    var dataSetX = configTable.DatTenCorporate(dataSet);

                                                                    string ColumOfDB = config[configs.sukiencol];

                                                                    string[] columns = ColumOfDB.Split(',');

                                                                    DataTable tableSK = new DataTable();
                                                                    foreach (string column in columns)
                                                                    {
                                                                        tableSK.Columns.Add(column);
                                                                    }
                                                                    DataTable dataTable = dataSetX.Tables["su kien"];
                                                                    string beginRow1 = config[configs.sukiencell];

                                                                    string[] column1 = beginRow1.Split(',');

                                                                    for (int y = Int32.Parse(config[configs.sukienrow]); y < dataTable.Rows.Count; y++)
                                                                    {
                                                                        ETableCorporate tblCorporate = new ETableCorporate();
                                                                        tblCorporate.CreateDate = DateTime.Now;
                                                                        tblCorporate.TransDate = group6Date;
                                                                        if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0" && dataTable.Rows[y][column1[0]].ToString().Length <= 15)
                                                                        {
                                                                            tblCorporate.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                            if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.SectionCode = view;
                                                                            }
                                                                            else tblCorporate.SectionCode = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.OutstandingShare = view;
                                                                            }
                                                                            else tblCorporate.OutstandingShare = 0;
                                                                            if (dataTable.Rows[y][column1[3]].ToString() != "")
                                                                            {
                                                                                tblCorporate.TypeOfAction = (string)dataTable.Rows[y][column1[3]];
                                                                            }
                                                                            else tblCorporate.TypeOfAction = "NULL";

                                                                            if (dataTable.Rows[y][column1[4]].ToString() != "")
                                                                            {
                                                                                tblCorporate.ExDate = dataTable.Rows[y][column1[4]].ToString();

                                                                            }
                                                                            else tblCorporate.ExDate = "NULL";

                                                                            if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.OfferPrice = view;
                                                                            }
                                                                            else tblCorporate.OfferPrice = 0;
                                                                            if (dataTable.Rows[y][column1[6]].ToString() != "")
                                                                            {
                                                                                tblCorporate.ExerciseRatio = dataTable.Rows[y][column1[6]].ToString();
                                                                            }
                                                                            else tblCorporate.ExerciseRatio = "NULL";
                                                                            if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.RatioForAdjustedPrice = view;
                                                                            }
                                                                            else tblCorporate.RatioForAdjustedPrice = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[8]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.PriorDayClose = view;
                                                                            }
                                                                            else tblCorporate.PriorDayClose = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[9]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.RefPriceofExDate = view;
                                                                            }
                                                                            else tblCorporate.RefPriceofExDate = 0;
                                                                            if (float.TryParse(dataTable.Rows[y][column1[10]].ToString(), out view))
                                                                            {
                                                                                tblCorporate.OutstandingShareAfterTheAdjustion = view;
                                                                            }
                                                                            else tblCorporate.OutstandingShareAfterTheAdjustion = 0;

                                                                            tableSK.Rows.Add(tblCorporate.TransDate, tblCorporate.CreateDate, tblCorporate.StockCode, tblCorporate.SectionCode, tblCorporate.OutstandingShare, tblCorporate.TypeOfAction, tblCorporate.ExDate, tblCorporate.OfferPrice, tblCorporate.ExerciseRatio, tblCorporate.RatioForAdjustedPrice, tblCorporate.PriorDayClose, tblCorporate.RefPriceofExDate, tblCorporate.OutstandingShareAfterTheAdjustion);
                                                                        }
                                                                    }
                                                                    StringBuilder sb = new StringBuilder();
                                                                    sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableCorporate]).Append(" ").Append(configs.valueTbl).Append(" ");
                                                                    foreach (DataRow row in tableSK.Rows)
                                                                    {
                                                                        sb.Append("(");
                                                                        sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                        sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                        sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                        sb.Append(row["SectionCode"]).Append(",");
                                                                        sb.Append(row["OutstandingShares"]).Append(",");
                                                                        sb.Append("'" + row["TypeOfAction"] + "'").Append(",");
                                                                        sb.Append("'" + row["ExDate"] + "'").Append(",");
                                                                        sb.Append(row["OfferPrice"]).Append(",");
                                                                        sb.Append("'" + row["ExerciseRatio"] + "'").Append(",");
                                                                        sb.Append(row["RatioForAdjustedPrice"]).Append(",");
                                                                        sb.Append(row["PriorDayClose"]).Append(",");
                                                                        sb.Append(row["RefPriceofExDate"]).Append(",");
                                                                        sb.Append(row["OutstandingShareAfterTheAdjustion"]);
                                                                        sb.Append("),");
                                                                    }
                                                                    if (tableSK.Rows.Count > 0)
                                                                    {
                                                                        command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                        command.ExecuteNonQuery();
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }

                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.Constituents))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");
                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();

                                                                    float view;
                                                                    string sheetName = config[configs.consSheetname];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {

                                                                            if (sheet == "VNAll")
                                                                            {
                                                                                DataSet dataSetX = new DataSet();

                                                                                if (group6Date >= startDate12 && group6Date <= endDate12 || (group6Date >= startDate13 && group6Date <= endDate13))
                                                                                {
                                                                                    dataSetX = configTable.DatTenVNAll1(dataSet);
                                                                                    string ColumOfDB1 = config[configs.conscol1];
                                                                                    string[] columns1 = ColumOfDB1.Split(',');

                                                                                    DataTable tableVNAII1 = new DataTable();
                                                                                    foreach (string column1 in columns1)
                                                                                    {
                                                                                        tableVNAII1.Columns.Add(column1);
                                                                                    }

                                                                                    DataTable dataTable = dataSetX.Tables["VNAll"];

                                                                                    string beginRow1 = config[configs.conscell1];
                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = Int32.Parse(config[configs.consrow]); y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableConstituents tblConstituents = new ETableConstituents();
                                                                                        tblConstituents.CreateDate = DateTime.Now;
                                                                                        tblConstituents.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                        {

                                                                                            tblConstituents.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.TodayClose = view;
                                                                                            }
                                                                                            else tblConstituents.TodayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.OutstandingShares = view;
                                                                                            }
                                                                                            else tblConstituents.OutstandingShares = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.FreeFloat = view;
                                                                                            }
                                                                                            else tblConstituents.FreeFloat = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.CapRatio = view;
                                                                                            }
                                                                                            else tblConstituents.CapRatio = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.FreeFloatAdjustedMarketCap = view;
                                                                                            }
                                                                                            else tblConstituents.FreeFloatAdjustedMarketCap = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.Weight = view;
                                                                                            }
                                                                                            else tblConstituents.Weight = 0;

                                                                                            tableVNAII1.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.TodayClose, tblConstituents.OutstandingShares, tblConstituents.FreeFloat, tblConstituents.CapRatio, tblConstituents.FreeFloatAdjustedMarketCap, tblConstituents.Weight);
                                                                                        }
                                                                                    }
                                                                                    StringBuilder sb = new StringBuilder();
                                                                                    sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableConstituents1]).Append(" (").Append(config[configs.conscol1]).Append(") ").Append(configs.valueTbl).Append(" ");
                                                                                    foreach (DataRow row in tableVNAII1.Rows)
                                                                                    {
                                                                                        sb.Append("(");
                                                                                        sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                                        sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                                        sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                                        sb.Append(row["TodayClose"]).Append(",");
                                                                                        sb.Append(row["OutstandingShares"]).Append(",");
                                                                                        sb.Append(row["FreeFloat"]).Append(",");
                                                                                        sb.Append(row["CapRatio"]).Append(",");
                                                                                        sb.Append(row["FreeFloatAdjustedMarketCap"]).Append(",");
                                                                                        sb.Append(row["Weight"]);
                                                                                        sb.Append("),");
                                                                                    }

                                                                                    if (tableVNAII1.Rows.Count > 0)
                                                                                    {
                                                                                        command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                                        command.ExecuteNonQuery();
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    dataSetX = configTable.DatTenVNAll(dataSet);
                                                                                    string ColumOfDB = config[configs.conscol];
                                                                                    string[] columns = ColumOfDB.Split(',');
                                                                                    DataTable tableVNAII = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tableVNAII.Columns.Add(column1);
                                                                                    }

                                                                                    DataTable dataTable = dataSetX.Tables["VNAll"];

                                                                                    string beginRow1 = config[configs.conscell];
                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = Int32.Parse(config[configs.consrow]); y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableConstituents tblConstituents = new ETableConstituents();
                                                                                        tblConstituents.CreateDate = DateTime.Now;
                                                                                        tblConstituents.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                        {

                                                                                            tblConstituents.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.TodayClose = view;
                                                                                            }
                                                                                            else tblConstituents.TodayClose = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.OutstandingShares = view;
                                                                                            }
                                                                                            else tblConstituents.OutstandingShares = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.ShareRestrictedOnTransfer = view;
                                                                                            }
                                                                                            else tblConstituents.ShareRestrictedOnTransfer = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.FreeFloat = view;
                                                                                            }
                                                                                            else tblConstituents.FreeFloat = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.CapRatio = view;
                                                                                            }
                                                                                            else tblConstituents.CapRatio = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.FreeFloatAdjustedMarketCap = view;
                                                                                            }
                                                                                            else tblConstituents.FreeFloatAdjustedMarketCap = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                            {
                                                                                                tblConstituents.Weight = view;
                                                                                            }
                                                                                            else tblConstituents.Weight = 0;

                                                                                            tableVNAII.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.TodayClose, tblConstituents.OutstandingShares, tblConstituents.ShareRestrictedOnTransfer, tblConstituents.FreeFloat, tblConstituents.CapRatio, tblConstituents.FreeFloatAdjustedMarketCap, tblConstituents.Weight);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tableVNAII, config[configs.tableConstituents1]);
                                                                                }
                                                                            }
                                                                            else if (sheet == "VN30")
                                                                            {
                                                                                var dataSetX = configTable.DatTenVN30(dataSet);

                                                                                string ColumOfDB = config[configs.conscolvn30];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table = new DataTable();
                                                                                foreach (string column in columns)
                                                                                {
                                                                                    table.Columns.Add(column);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["VN30"];
                                                                                string beginRow1 = config[configs.conscellvn30];

                                                                                string[] column1 = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.consrowvn30]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableConstituents tblConstituents = new ETableConstituents();
                                                                                    tblConstituents.CreateDate = DateTime.Now;
                                                                                    tblConstituents.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                    {
                                                                                        tblConstituents.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                                        if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.PriceClose = view;
                                                                                        }
                                                                                        else tblConstituents.PriceClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.OutstandingShares = view;
                                                                                        }
                                                                                        else tblConstituents.OutstandingShares = 0;
                                                                                        if (sheet == "VN30")
                                                                                        {
                                                                                            tblConstituents.Type = "VN30";
                                                                                        }

                                                                                        table.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.PriceClose, tblConstituents.OutstandingShares, tblConstituents.Type);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table, config[configs.tableConstituents2]);
                                                                            }
                                                                            else if (sheet == "VNMIDCAP")
                                                                            {
                                                                                var dataSetX = configTable.DatTenVNMIDCAP(dataSet);

                                                                                string ColumOfDB = config[configs.conscolVNMIDCAP];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table = new DataTable();
                                                                                foreach (string column in columns)
                                                                                {
                                                                                    table.Columns.Add(column);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["VNMIDCAP"];
                                                                                string beginRow1 = config[configs.conscellVNMIDCAP];

                                                                                string[] column1 = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.consrowVNMIDCAP]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableConstituents tblConstituents = new ETableConstituents();
                                                                                    tblConstituents.CreateDate = DateTime.Now;
                                                                                    tblConstituents.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                    {
                                                                                        tblConstituents.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                                        if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.PriceClose = view;
                                                                                        }
                                                                                        else tblConstituents.PriceClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.OutstandingShares = view;
                                                                                        }
                                                                                        else tblConstituents.OutstandingShares = 0;

                                                                                        tblConstituents.Type = "VNMIDCAP";
                                                                                        table.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.PriceClose, tblConstituents.OutstandingShares, tblConstituents.Type);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table, config[configs.tableConstituents2]);

                                                                            }
                                                                            else if (sheet == "VNSMALLCAP")
                                                                            {
                                                                                var dataSetX = configTable.DatTenVNSMALLCAP(dataSet);

                                                                                string ColumOfDB = config[configs.conscolVNSMALLCAP];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table = new DataTable();
                                                                                foreach (string column in columns)
                                                                                {
                                                                                    table.Columns.Add(column);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["VNSMALLCAP"];
                                                                                string beginRow1 = config[configs.conscellVNSMALLCAP];

                                                                                string[] column1 = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.consrowVNSMALLCAP]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableConstituents tblConstituents = new ETableConstituents();
                                                                                    tblConstituents.CreateDate = DateTime.Now;
                                                                                    tblConstituents.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                    {
                                                                                        tblConstituents.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                                        if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.PriceClose = view;
                                                                                        }
                                                                                        else tblConstituents.PriceClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.OutstandingShares = view;
                                                                                        }
                                                                                        else tblConstituents.OutstandingShares = 0;

                                                                                        tblConstituents.Type = "VNSMALLCAP";
                                                                                        table.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.PriceClose, tblConstituents.OutstandingShares, tblConstituents.Type);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table, config[configs.tableConstituents2]);

                                                                            }
                                                                            else if (sheet == "VN100")
                                                                            {
                                                                                var dataSetX = configTable.DatTenVN100(dataSet);

                                                                                string ColumOfDB = config[configs.conscolVN100];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table = new DataTable();
                                                                                foreach (string column in columns)
                                                                                {
                                                                                    table.Columns.Add(column);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["VN100"];
                                                                                string beginRow1 = config[configs.conscellVN100];

                                                                                string[] column1 = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.consrowVN100]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableConstituents tblConstituents = new ETableConstituents();
                                                                                    tblConstituents.CreateDate = DateTime.Now;
                                                                                    tblConstituents.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                    {
                                                                                        tblConstituents.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                                        if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.PriceClose = view;
                                                                                        }
                                                                                        else tblConstituents.PriceClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.OutstandingShares = view;
                                                                                        }
                                                                                        else tblConstituents.OutstandingShares = 0;

                                                                                        tblConstituents.Type = "VN100";
                                                                                        table.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.PriceClose, tblConstituents.OutstandingShares, tblConstituents.Type);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table, config[configs.tableConstituents2]);

                                                                            }
                                                                            else if (sheet == "VNALLSHARE")
                                                                            {
                                                                                var dataSetX = configTable.DatTenVNALLSHARE(dataSet);

                                                                                string ColumOfDB = config[configs.conscolVNALLSHARE];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table = new DataTable();
                                                                                foreach (string column in columns)
                                                                                {
                                                                                    table.Columns.Add(column);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["VNALLSHARE"];
                                                                                string beginRow1 = config[configs.conscellVNALLSHARE];

                                                                                string[] column1 = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.consrowVNALLSHARE]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableConstituents tblConstituents = new ETableConstituents();
                                                                                    tblConstituents.CreateDate = DateTime.Now;
                                                                                    tblConstituents.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                    {
                                                                                        tblConstituents.StockCode = (string)dataTable.Rows[y][column1[0]];

                                                                                        if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.PriceClose = view;
                                                                                        }
                                                                                        else tblConstituents.PriceClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                        {
                                                                                            tblConstituents.OutstandingShares = view;
                                                                                        }
                                                                                        else tblConstituents.OutstandingShares = 0;

                                                                                        tblConstituents.Type = "VNALLSHARE";
                                                                                        table.Rows.Add(tblConstituents.TransDate, tblConstituents.CreateDate, tblConstituents.StockCode, tblConstituents.PriceClose, tblConstituents.OutstandingShares, tblConstituents.Type);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table, config[configs.tableConstituents2]);

                                                                            }
                                                                            else if (sheet == "HOSEINDEX")
                                                                            {
                                                                                var dataSetX = configTable.DatTenHOSEINDEX(dataSet);

                                                                                string ColumOfDB = config[configs.data1col11];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableHOSEINDEX = new DataTable();
                                                                                foreach (string column in columns)
                                                                                {
                                                                                    tableHOSEINDEX.Columns.Add(column);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["HOSEINDEX"];
                                                                                string beginRow1 = config[configs.databeginCell11];

                                                                                string[] column1 = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.databeginRow11]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETable tbl = new ETable();
                                                                                    tbl.CreateDate = DateTime.Now;
                                                                                    tbl.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column1[0]].ToString() != "" && dataTable.Rows[y][column1[0]].ToString() != "0")
                                                                                    {
                                                                                        tbl.IndexName = (string)dataTable.Rows[y][column1[0]];

                                                                                        if (float.TryParse(dataTable.Rows[y][column1[1]].ToString(), out view))
                                                                                        {
                                                                                            tbl.OpenIndexValue = view;
                                                                                        }
                                                                                        else tbl.OpenIndexValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[2]].ToString(), out view))
                                                                                        {
                                                                                            tbl.CloseIndexValue = view;
                                                                                        }
                                                                                        else tbl.CloseIndexValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[3]].ToString(), out view))
                                                                                        {
                                                                                            tbl.High = view;
                                                                                        }
                                                                                        else tbl.High = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[4]].ToString(), out view))
                                                                                        {
                                                                                            tbl.Low = view;
                                                                                        }
                                                                                        else tbl.Low = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[5]].ToString(), out view))
                                                                                        {
                                                                                            tbl.UpDown = view;
                                                                                        }
                                                                                        else tbl.UpDown = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[6]].ToString(), out view))
                                                                                        {
                                                                                            tbl.Change = view;
                                                                                        }
                                                                                        else tbl.Change = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[7]].ToString(), out view))
                                                                                        {
                                                                                            tbl.TradingVolume = view;
                                                                                        }
                                                                                        else tbl.TradingVolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[8]].ToString(), out view))
                                                                                        {
                                                                                            tbl.TradingValue = view;
                                                                                        }
                                                                                        else tbl.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column1[9]].ToString(), out view))
                                                                                        {
                                                                                            tbl.Marketcap = view;
                                                                                        }
                                                                                        else tbl.Marketcap = 0;


                                                                                        tableHOSEINDEX.Rows.Add(tbl.TransDate, tbl.CreateDate, tbl.IndexName, tbl.OpenIndexValue, tbl.CloseIndexValue, tbl.High, tbl.Low, tbl.UpDown, tbl.Change, tbl.TradingVolume, tbl.TradingValue, tbl.Marketcap);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableHOSEINDEX, config[configs.tableSessionHOSE]);

                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.Order1) || group7Value.Contains(configs.Order2) || group7Value.Contains(configs.Order3))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();
                                                                    float view;
                                                                    string sheetName = config[configs.data3Sheetname];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {
                                                                            if (sheet == "Chi tiet (Details)")
                                                                            {
                                                                                var dataSetX = configTable.DatTenOrderChitiet(dataSet);
                                                                                string ColumOfDB = config[configs.data3col];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table1chitiet = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    table1chitiet.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["Chi tiet (Details)"];

                                                                                string beginRow1 = config[configs.data3cell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.data3row]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableOrder tblOrder = new ETableOrder();
                                                                                    tblOrder.CreateDate = DateTime.Now;
                                                                                    tblOrder.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblOrder.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuyingOrders = view;
                                                                                        }
                                                                                        else tblOrder.BuyingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuyingVolume = view;
                                                                                        }
                                                                                        else tblOrder.BuyingVolume = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.SellingOrders = view;
                                                                                        }
                                                                                        else tblOrder.SellingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.SellingVolume = view;
                                                                                        }
                                                                                        else tblOrder.SellingVolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.TradingVolume = view;
                                                                                        }
                                                                                        else tblOrder.TradingVolume = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuySellVolume = view;
                                                                                        }
                                                                                        else tblOrder.BuySellVolume = 0;


                                                                                        table1chitiet.Rows.Add(tblOrder.TransDate, tblOrder.CreateDate, tblOrder.StockCode, tblOrder.BuyingOrders, tblOrder.BuyingVolume, tblOrder.SellingOrders, tblOrder.SellingVolume, tblOrder.TradingVolume, tblOrder.BuySellVolume);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table1chitiet, config[configs.tableOrder1]);
                                                                            }
                                                                            else if (sheet == "CW")
                                                                            {
                                                                                var dataSetX = configTable.DatTenOrderCW(dataSet);
                                                                                string ColumOfDB = config[configs.data2col];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table2cw = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    table2cw.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["CW"];

                                                                                string beginRow1 = config[configs.data2cell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.data3row]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableOrder tblOrder = new ETableOrder();
                                                                                    tblOrder.CreateDate = DateTime.Now;
                                                                                    tblOrder.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblOrder.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuyingOrders = view;
                                                                                        }
                                                                                        else tblOrder.BuyingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuyingVolume = view;
                                                                                        }
                                                                                        else tblOrder.BuyingVolume = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.SellingOrders = view;
                                                                                        }
                                                                                        else tblOrder.SellingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.SellingVolume = view;
                                                                                        }
                                                                                        else tblOrder.SellingVolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.TradingVolume = view;
                                                                                        }
                                                                                        else tblOrder.TradingVolume = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuySellVolume = view;
                                                                                        }
                                                                                        else tblOrder.BuySellVolume = 0;

                                                                                        table2cw.Rows.Add(tblOrder.TransDate, tblOrder.CreateDate, tblOrder.StockCode, tblOrder.BuyingOrders, tblOrder.BuyingVolume, tblOrder.SellingOrders, tblOrder.SellingVolume, tblOrder.TradingVolume, tblOrder.BuySellVolume);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table2cw, config[configs.tableOrderCW]);

                                                                            }
                                                                            else if (sheet == "ODD")
                                                                            {
                                                                                var dataSetX = configTable.DatTenOrderODD(dataSet);
                                                                                string ColumOfDB = config[configs.data2col];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable table3odd = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    table3odd.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["ODD"];

                                                                                string beginRow1 = config[configs.data2cell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.data3row]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableOrder tblOrder = new ETableOrder();
                                                                                    tblOrder.CreateDate = DateTime.Now;
                                                                                    tblOrder.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblOrder.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuyingOrders = view;
                                                                                        }
                                                                                        else tblOrder.BuyingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuyingVolume = view;
                                                                                        }
                                                                                        else tblOrder.BuyingVolume = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.SellingOrders = view;
                                                                                        }
                                                                                        else tblOrder.SellingOrders = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.SellingVolume = view;
                                                                                        }
                                                                                        else tblOrder.SellingVolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.TradingVolume = view;
                                                                                        }
                                                                                        else tblOrder.TradingVolume = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblOrder.BuySellVolume = view;
                                                                                        }
                                                                                        else tblOrder.BuySellVolume = 0;

                                                                                        table3odd.Rows.Add(tblOrder.TransDate, tblOrder.CreateDate, tblOrder.StockCode, tblOrder.BuyingOrders, tblOrder.BuyingVolume, tblOrder.SellingOrders, tblOrder.SellingVolume, tblOrder.TradingVolume, tblOrder.BuySellVolume);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(table3odd, config[configs.tableOrderODD]);

                                                                            }

                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }

                                                    else if (group7Value.Contains(configs.Foreign2) || group7Value.Contains(configs.Foreign1) || group7Value.Contains(configs.Foreign4))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");
                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();
                                                                    float view;

                                                                    string sheetName = config[configs.foreignsheetname];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {
                                                                            if (sheet == "1")
                                                                            {
                                                                                var dataSetX = configTable.DatTenForeign1(dataSet);
                                                                                string ColumOfDB = config[configs.foreig1ncol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeign1 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeign1.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["1"];

                                                                                string beginRow1 = config[configs.foreign1cell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.foreignrow]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;
                                                                                    if (y == 10)
                                                                                    {
                                                                                        tblForeign.Type = "OrderBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 11)
                                                                                    {
                                                                                        tblForeign.Type = "OrderSelling";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 12)
                                                                                    {
                                                                                        tblForeign.Type = "OrderDifference";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 14)
                                                                                    {
                                                                                        tblForeign.Type = "PutBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 15)
                                                                                    {
                                                                                        tblForeign.Type = "PutSelling";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 16)
                                                                                    {
                                                                                        tblForeign.Type = "PutDifference";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 18)
                                                                                    {
                                                                                        tblForeign.Type = "TotalBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 19)
                                                                                    {
                                                                                        tblForeign.Type = "TotalBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 20)
                                                                                    {
                                                                                        tblForeign.Type = "TotalDifference";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.Tradingvolume = view;
                                                                                        }
                                                                                        else tblForeign.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblForeign.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TradingValue = view;
                                                                                        }
                                                                                        else tblForeign.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblForeign.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else continue;

                                                                                    tableforeign1.Rows.Add(tblForeign.TransDate, tblForeign.CreateDate, tblForeign.Type, tblForeign.Tradingvolume, tblForeign.rateEntireMaket, tblForeign.TradingValue, tblForeign.ratEntireMaket2);
                                                                                }

                                                                                StringBuilder sb = new StringBuilder();
                                                                                sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableForeign1]).Append(" (").Append(config[configs.foreig1ncol]).Append(") ").Append(configs.valueTbl).Append(" ");
                                                                                foreach (DataRow row in tableforeign1.Rows)
                                                                                {
                                                                                    sb.Append("(");
                                                                                    sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                                    sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                                    sb.Append("'" + row["Type"] + "'").Append(",");
                                                                                    sb.Append(row["Tradingvolume"]).Append(",");
                                                                                    sb.Append(row["rateEntireMaket"]).Append(",");
                                                                                    sb.Append(row["TradingValue"]).Append(",");
                                                                                    sb.Append(row["ratEntireMaket2"]);
                                                                                    sb.Append("),");
                                                                                }
                                                                                if (tableforeign1.Rows.Count > 0)
                                                                                {
                                                                                    command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                                    command.ExecuteNonQuery();
                                                                                }
                                                                            }
                                                                            else if (sheet == "2")
                                                                            {
                                                                                DataSet dataSetX = new DataSet();
                                                                                string beginRow1 = "";
                                                                                if (group7Value.Contains(configs.Foreign1))
                                                                                {
                                                                                    dataSetX = configTable.DatTenForeign2(dataSet);
                                                                                    beginRow1 = config[configs.foreign2cell2];
                                                                                }
                                                                                else if (group7Value.Contains(configs.Foreign2)) ;
                                                                                {
                                                                                    dataSetX = configTable.DatTenForeign2cu(dataSet);
                                                                                    beginRow1 = config[configs.foreign2cell1];
                                                                                }
                                                                                string ColumOfDB = config[configs.foreig2ncol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeign2 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeign2.Columns.Add(column1);
                                                                                }

                                                                                DataTable dataTable = dataSetX.Tables["2"];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.foreig2nrow]); y < dataTable.Rows.Count - 4; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TotalRoom = view;
                                                                                        }
                                                                                        else tblForeign.TotalRoom = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.CurrentRoom = view;
                                                                                        }
                                                                                        else tblForeign.CurrentRoom = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ForeignOwnedRatio = view;
                                                                                        }
                                                                                        else tblForeign.ForeignOwnedRatio = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.StateOwnedRatio = view;
                                                                                        }
                                                                                        else tblForeign.StateOwnedRatio = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVolPreOpen = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolPreOpen = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVolCount = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolCount = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVolPreClose = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolPreClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyValue = view;
                                                                                        }
                                                                                        else tblForeign.BuyValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVolPreOpen = view;
                                                                                        }
                                                                                        else tblForeign.SellVolPreOpen = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVolCount = view;
                                                                                        }
                                                                                        else tblForeign.SellVolCount = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVolPreClose = view;
                                                                                        }
                                                                                        else tblForeign.SellVolPreClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellValue = view;
                                                                                        }
                                                                                        else tblForeign.SellValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellValue = view;
                                                                                        }
                                                                                        else tblForeign.SellValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVol = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVal = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutSellVol = view;
                                                                                        }
                                                                                        else tblForeign.PutSellVol = 0;
                                                                                        tableforeign2.Rows.Add(tblForeign.TransDate, tblForeign.CreateDate, tblForeign.StockCode, tblForeign.TotalRoom, tblForeign.CurrentRoom, tblForeign.ForeignOwnedRatio, tblForeign.StateOwnedRatio, tblForeign.BuyVolPreOpen, tblForeign.BuyVolCount, tblForeign.BuyVolPreClose, tblForeign.BuyValue, tblForeign.SellVolPreOpen, tblForeign.SellVolCount, tblForeign.SellVolPreClose, tblForeign.SellValue, tblForeign.PutBuyVol, tblForeign.PutBuyVal, tblForeign.PutSellVol, tblForeign.PutSellVal);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableforeign2, config[configs.tableForeign2]);


                                                                            }
                                                                            else if (sheet == "3")
                                                                            {
                                                                                var dataSetX = configTable.DatTenForeign3(dataSet);
                                                                                string ColumOfDB = config[configs.foreig3ncol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeign3 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeign3.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["3"];

                                                                                string beginRow1 = config[configs.foreign3cell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.foreign3row]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                    {
                                                                                        tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVol = view;
                                                                                        }
                                                                                        else tblForeign.BuyVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyValue = view;
                                                                                        }
                                                                                        else tblForeign.BuyValue = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVol = view;
                                                                                        }
                                                                                        else tblForeign.SellVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellValue = view;
                                                                                        }
                                                                                        else tblForeign.SellValue = 0;
                                                                                        tableforeign3.Rows.Add(tblForeign.TransDate, tblForeign.CreateDate, tblForeign.StockCode, tblForeign.BuyVol, tblForeign.BuyValue, tblForeign.SellVol, tblForeign.SellValue);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableforeign3, config[configs.tableForeign3]);

                                                                            }
                                                                            else if (sheet == "5")
                                                                            {
                                                                                var dataSetX = configTable.DatTenForeign5(dataSet);
                                                                                string ColumOfDB = config[configs.foreign5col];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeign5 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeign5.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["5"];

                                                                                string beginRow1 = config[configs.foreign5cell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.foreign5row]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString().Length <= 20)
                                                                                    {
                                                                                        tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TotalRoom = view;
                                                                                        }
                                                                                        else tblForeign.TotalRoom = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.CurrentRoom = view;
                                                                                        }
                                                                                        else tblForeign.CurrentRoom = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ForeignOwnedRatio = view;
                                                                                        }
                                                                                        else tblForeign.ForeignOwnedRatio = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.StateOwnedRatio = view;
                                                                                        }
                                                                                        else tblForeign.StateOwnedRatio = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderBuyPreOpen = view;
                                                                                        }
                                                                                        else tblForeign.OrderBuyPreOpen = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderBuyCont = view;
                                                                                        }
                                                                                        else tblForeign.OrderBuyCont = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderBuyPreClose = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolPreClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderSellPreOpen = view;
                                                                                        }
                                                                                        else tblForeign.OrderSellPreOpen = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OpenSellCont = view;
                                                                                        }
                                                                                        else tblForeign.OpenSellCont = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OpenSellClose = view;
                                                                                        }
                                                                                        else tblForeign.OpenSellClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVol = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutSellVol = view;
                                                                                        }
                                                                                        else tblForeign.PutSellVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TotalBuyVol = view;
                                                                                        }
                                                                                        else tblForeign.TotalBuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TotalSellVol = view;
                                                                                        }
                                                                                        else tblForeign.TotalSellVol = 0;

                                                                                        tableforeign5.Rows.Add(tblForeign.CreateDate, tblForeign.TransDate, tblForeign.StockCode, tblForeign.TotalRoom, tblForeign.CurrentRoom, tblForeign.ForeignOwnedRatio, tblForeign.StateOwnedRatio, tblForeign.OrderBuyPreOpen, tblForeign.OrderBuyCont, tblForeign.OrderBuyPreClose, tblForeign.OrderSellPreOpen, tblForeign.OpenSellCont, tblForeign.OpenSellClose, tblForeign.PutBuyVol, tblForeign.PutSellVol, tblForeign.TotalBuyVol, tblForeign.TotalSellVol);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableforeign5, config[configs.tableForeign5]);
                                                                            }
                                                                            else if (sheet == "CW")
                                                                            {
                                                                                var dataSetX = configTable.DatTenForeignCW(dataSet);
                                                                                string ColumOfDB = config[configs.foreigncwcol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeignCW = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeignCW.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["CW"];

                                                                                string beginRow1 = config[configs.foreigncwcell];

                                                                                string[] column = beginRow1.Split(',');
                                                                                int numberRow = dataTable.Rows.Count;
                                                                                for (int y = Int32.Parse(config[configs.foreigncwrow]); y < numberRow; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString().Length <= 20)
                                                                                    {
                                                                                        tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.TotalRoom = view;
                                                                                        }
                                                                                        else tblForeign.TotalRoom = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.CurrentRoom = view;
                                                                                        }
                                                                                        else tblForeign.CurrentRoom = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.ForeignOwnedRatio = view;
                                                                                        }
                                                                                        else tblForeign.ForeignOwnedRatio = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.StateOwnedRatio = view;
                                                                                        }
                                                                                        else tblForeign.StateOwnedRatio = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVolPreOpen = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolPreOpen = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVolCount = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolCount = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVolPreClose = view;
                                                                                        }
                                                                                        else tblForeign.BuyVolPreClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyValue = view;
                                                                                        }
                                                                                        else tblForeign.BuyValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVolPreOpen = view;
                                                                                        }
                                                                                        else tblForeign.SellVolPreOpen = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVolCount = view;
                                                                                        }
                                                                                        else tblForeign.SellVolCount = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVolPreClose = view;
                                                                                        }
                                                                                        else tblForeign.SellVolPreClose = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellValue = view;
                                                                                        }
                                                                                        else tblForeign.SellValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellValue = view;
                                                                                        }
                                                                                        else tblForeign.SellValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVol = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVal = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutSellVol = view;
                                                                                        }
                                                                                        else tblForeign.PutSellVol = 0;
                                                                                        tableforeignCW.Rows.Add(tblForeign.TransDate, tblForeign.CreateDate, tblForeign.StockCode, tblForeign.TotalRoom, tblForeign.CurrentRoom, tblForeign.ForeignOwnedRatio, tblForeign.StateOwnedRatio, tblForeign.BuyVolPreOpen, tblForeign.BuyVolCount, tblForeign.BuyVolPreClose, tblForeign.BuyValue, tblForeign.SellVolPreOpen, tblForeign.SellVolCount, tblForeign.SellVolPreClose, tblForeign.SellValue, tblForeign.PutBuyVol, tblForeign.PutBuyVal, tblForeign.PutSellVol, tblForeign.PutSellVal);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableforeignCW, config[configs.tableForeignCW]);

                                                                            }
                                                                            else if (sheet == "ODD")
                                                                            {
                                                                                var dataSetX = configTable.DatTenForeignODD(dataSet);
                                                                                string ColumOfDB = config[configs.foreignoddcol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeignodd = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeignodd.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["ODD"];

                                                                                string beginRow1 = config[configs.foreignoddcell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.foreignoddrow]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString().Length <= 20)
                                                                                    {
                                                                                        tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderBuyVol = view;
                                                                                        }
                                                                                        else tblForeign.OrderBuyVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderSellVol = view;
                                                                                        }
                                                                                        else tblForeign.OrderSellVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderBuyVal = view;
                                                                                        }
                                                                                        else tblForeign.OrderBuyVal = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.OrderSellVal = view;
                                                                                        }
                                                                                        else tblForeign.OrderSellVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVol = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutSellVol = view;
                                                                                        }
                                                                                        else tblForeign.PutSellVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutBuyVal = view;
                                                                                        }
                                                                                        else tblForeign.PutBuyVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.PutSellVal = view;
                                                                                        }
                                                                                        else tblForeign.PutSellVal = 0;

                                                                                        tableforeignodd.Rows.Add(tblForeign.TransDate, tblForeign.CreateDate, tblForeign.StockCode, tblForeign.OrderBuyVol, tblForeign.OrderSellVol, tblForeign.OrderBuyVal, tblForeign.OrderSellVal, tblForeign.PutBuyVol, tblForeign.PutSellVol, tblForeign.PutBuyVal, tblForeign.PutSellVal);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableforeignodd, config[configs.tableForeignODD]);
                                                                            }
                                                                            else if (sheet == "6")
                                                                            {
                                                                                var dataSetX = configTable.DatTenForeign6(dataSet);
                                                                                string ColumOfDB = config[configs.foreign6col];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableforeign3 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableforeign3.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["6"];

                                                                                string beginRow1 = config[configs.foreign6cell];

                                                                                string[] column = beginRow1.Split(',');
                                                                                int numberRow = dataTable.Rows.Count;
                                                                                for (int y = Int32.Parse(config[configs.foreign6row]); y < numberRow; y++)
                                                                                {
                                                                                    ETableForeign tblForeign = new ETableForeign();
                                                                                    tblForeign.CreateDate = DateTime.Now;
                                                                                    tblForeign.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                    {
                                                                                        tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.BuyVol = view;
                                                                                        }
                                                                                        else tblForeign.BuyVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblForeign.SellVol = view;
                                                                                        }
                                                                                        else tblForeign.SellVol = 0;

                                                                                        tableforeign3.Rows.Add(tblForeign.TransDate, tblForeign.CreateDate, tblForeign.StockCode, tblForeign.BuyVol, tblForeign.SellVol);
                                                                                    }
                                                                                }
                                                                                StringBuilder sb = new StringBuilder();
                                                                                sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableForeign3]).Append(" (").Append(config[configs.foreign6col]).Append(") ").Append(configs.valueTbl).Append(" ");
                                                                                foreach (DataRow row in tableforeign3.Rows)
                                                                                {
                                                                                    sb.Append("(");
                                                                                    sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                                    sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                                    sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                                    sb.Append(row["BuyVol"]).Append(",");
                                                                                    sb.Append(row["SellVol"]);
                                                                                    sb.Append("),");
                                                                                }
                                                                                if (tableforeign3.Rows.Count > 0)
                                                                                {
                                                                                    command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                                    command.ExecuteNonQuery();
                                                                                }

                                                                            }
                                                                        }

                                                                    }

                                                                }

                                                            }

                                                            break;
                                                        }

                                                    }
                                                    else if (group7Value.Contains(configs.Proprietary))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");
                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();
                                                                    float view;

                                                                    string sheetName = config[configs.ProprietarySheetName];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {
                                                                            if (sheet == "Tong hop (Summary)")
                                                                            {
                                                                                var dataSetX = configTable.DatTenProprietarySummary(dataSet);
                                                                                string ColumOfDB = config[configs.ProprietaryCol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableproprietary1 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableproprietary1.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["Tong hop (Summary)"];

                                                                                string beginRow1 = config[configs.ProprietaryCell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.ProprietaryRow]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETableProprietary tblProprietary = new ETableProprietary();
                                                                                    tblProprietary.CreateDate = DateTime.Now;
                                                                                    tblProprietary.TransDate = group6Date;
                                                                                    if (y == 10)
                                                                                    {
                                                                                        tblProprietary.Type = "OrderBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 11)
                                                                                    {
                                                                                        tblProprietary.Type = "OrderSelling";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 12)
                                                                                    {
                                                                                        tblProprietary.Type = "OrderDifference";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 14)
                                                                                    {
                                                                                        tblProprietary.Type = "PutBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 15)
                                                                                    {
                                                                                        tblProprietary.Type = "PutSelling";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 16)
                                                                                    {
                                                                                        tblProprietary.Type = "PutDifference";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 18)
                                                                                    {
                                                                                        tblProprietary.Type = "TotalBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 19)
                                                                                    {
                                                                                        tblProprietary.Type = "TotalBuying";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else if (y == 20)
                                                                                    {
                                                                                        tblProprietary.Type = "TotalDifference";
                                                                                        if (float.TryParse(dataTable.Rows[y][column[0]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.Tradingvolume = view;
                                                                                        }
                                                                                        else tblProprietary.Tradingvolume = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.rateEntireMaket = view;
                                                                                        }
                                                                                        else tblProprietary.rateEntireMaket = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.TradingValue = view;
                                                                                        }
                                                                                        else tblProprietary.TradingValue = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.ratEntireMaket2 = view;
                                                                                        }
                                                                                        else tblProprietary.ratEntireMaket2 = 0;
                                                                                    }
                                                                                    else continue;

                                                                                    tableproprietary1.Rows.Add(tblProprietary.TransDate, tblProprietary.CreateDate, tblProprietary.Type, tblProprietary.Tradingvolume, tblProprietary.rateEntireMaket, tblProprietary.TradingValue, tblProprietary.ratEntireMaket2);

                                                                                }

                                                                                StringBuilder sb = new StringBuilder();
                                                                                sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableProprietary1]).Append(" (").Append(config[configs.ProprietaryCol]).Append(") ").Append(configs.valueTbl);
                                                                                foreach (DataRow row in tableproprietary1.Rows)
                                                                                {
                                                                                    sb.Append("(");
                                                                                    sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                                    sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                                    sb.Append("'" + row["Type"] + "'").Append(",");
                                                                                    sb.Append(row["Tradingvolume"]).Append(",");
                                                                                    sb.Append(row["rateEntireMaket"]).Append(",");
                                                                                    sb.Append(row["TradingValue"]).Append(",");
                                                                                    sb.Append(row["ratEntireMaket2"]);
                                                                                    sb.Append("),");
                                                                                }
                                                                                if (tableproprietary1.Rows.Count > 0)
                                                                                {
                                                                                    command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                                    command.ExecuteNonQuery();
                                                                                }

                                                                            }
                                                                            else if (sheet == "Chi tiet (Details)")
                                                                            {
                                                                                var dataSetX = configTable.DatTenProprietaryDetails(dataSet);
                                                                                string ColumOfDB = config[configs.Proprietary_DetailsyCol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableproprietary2 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableproprietary2.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["Chi tiet (Details)"];

                                                                                string beginRow1 = config[configs.Proprietary_DetailsCell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.Proprietary_DetailsRow]); y < dataTable.Rows.Count - 3; y++)
                                                                                {
                                                                                    ETableProprietary tblProprietary = new ETableProprietary();

                                                                                    tblProprietary.CreateDate = DateTime.Now;
                                                                                    tblProprietary.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblProprietary.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyRateVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellRateVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyVal = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyRateVal = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellRateVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyRateVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellRateVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyRateVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellRateVal = 0;

                                                                                        tableproprietary2.Rows.Add(tblProprietary.TransDate, tblProprietary.CreateDate, tblProprietary.StockCode, tblProprietary.OrderBuyVol, tblProprietary.OrderBuyRateVol, tblProprietary.OrderSellVol, tblProprietary.OrderSellRateVol, tblProprietary.OrderBuyVal, tblProprietary.OrderBuyRateVal, tblProprietary.OrderSellVal, tblProprietary.OrderSellRateVal, tblProprietary.PutBuyVol, tblProprietary.PutBuyRateVol, tblProprietary.PutSellVol, tblProprietary.PutSellRateVol, tblProprietary.PutBuyVal, tblProprietary.PutBuyRateVal, tblProprietary.PutSellVal, tblProprietary.PutSellRateVal);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableproprietary2, config[configs.tableProprietary2]);
                                                                            }
                                                                            else if (sheet == "Chi tiet (Details) CW")
                                                                            {
                                                                                var dataSetX = configTable.DatTenProprietaryDetailsCW(dataSet);
                                                                                string ColumOfDB = config[configs.Proprietary_DetailsCWCol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableproprietary3 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableproprietary3.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["Chi tiet (Details) CW"];

                                                                                string beginRow1 = config[configs.Proprietary_DetailsCWCell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.Proprietary_DetailsCWRow]); y < dataTable.Rows.Count - 3; y++)
                                                                                {
                                                                                    ETableProprietary tblProprietary = new ETableProprietary();

                                                                                    tblProprietary.CreateDate = DateTime.Now;
                                                                                    tblProprietary.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblProprietary.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyRateVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellVol = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellRateVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyVal = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderBuyRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderBuyRateVal = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.OrderSellRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.OrderSellRateVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyRateVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellRateVol = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellRateVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutBuyRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutBuyRateVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellVal = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.PutSellRateVal = view;
                                                                                        }
                                                                                        else tblProprietary.PutSellRateVal = 0;

                                                                                        tableproprietary3.Rows.Add(tblProprietary.TransDate, tblProprietary.CreateDate, tblProprietary.StockCode, tblProprietary.OrderBuyVol, tblProprietary.OrderBuyRateVol, tblProprietary.OrderSellVol, tblProprietary.OrderSellRateVol, tblProprietary.OrderBuyVal, tblProprietary.OrderBuyRateVal, tblProprietary.OrderSellVal, tblProprietary.OrderSellRateVal, tblProprietary.PutBuyVol, tblProprietary.PutBuyRateVol, tblProprietary.PutSellVol, tblProprietary.PutSellRateVol, tblProprietary.PutBuyVal, tblProprietary.PutBuyRateVal, tblProprietary.PutSellVal, tblProprietary.PutSellRateVal);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableproprietary3, config[configs.tableProprietary3]);
                                                                            }
                                                                            else if (sheet == "Dat lenh (Order)")
                                                                            {
                                                                                var dataSetX = configTable.DatTenProprietaryOrder(dataSet);
                                                                                string ColumOfDB = config[configs.Proprietary_OrderCol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableproprietary4 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableproprietary4.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["Dat lenh (Order)"];

                                                                                string beginRow1 = config[configs.Proprietary_OrderCell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.Proprietary_OrderRow]); y < dataTable.Rows.Count - 3; y++)
                                                                                {
                                                                                    ETableProprietary tblProprietary = new ETableProprietary();

                                                                                    tblProprietary.CreateDate = DateTime.Now;
                                                                                    tblProprietary.TransDate = group6Date;
                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblProprietary.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.BuyOrder = view;
                                                                                        }
                                                                                        else tblProprietary.BuyOrder = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.BuyVol = view;
                                                                                        }
                                                                                        else tblProprietary.BuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.SellOrder = view;
                                                                                        }
                                                                                        else tblProprietary.SellOrder = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.BuySellVol = view;
                                                                                        }
                                                                                        else tblProprietary.BuySellVol = 0;
                                                                                        tableproprietary4.Rows.Add(tblProprietary.TransDate, tblProprietary.CreateDate, tblProprietary.StockCode, tblProprietary.BuyOrder, tblProprietary.BuyVol, tblProprietary.SellOrder, tblProprietary.SellVol, tblProprietary.BuySellVol);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableproprietary4, config[configs.tableProprietary4]);

                                                                            }
                                                                            else if (sheet == "Dat lenh (Order) CW")
                                                                            {
                                                                                var dataSetX = configTable.DatTenProprietaryOrderCW(dataSet);
                                                                                string ColumOfDB = config[configs.Proprietary_OrderCWCol];

                                                                                string[] columns = ColumOfDB.Split(',');

                                                                                DataTable tableproprietary5 = new DataTable();
                                                                                foreach (string column1 in columns)
                                                                                {
                                                                                    tableproprietary5.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["Dat lenh (Order) CW"];

                                                                                string beginRow1 = config[configs.Proprietary_OrderCWCell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                for (int y = Int32.Parse(config[configs.Proprietary_OrderCWRow]); y < dataTable.Rows.Count - 3; y++)
                                                                                {
                                                                                    ETableProprietary tblProprietary = new ETableProprietary();

                                                                                    tblProprietary.CreateDate = DateTime.Now;
                                                                                    tblProprietary.TransDate = group6Date;

                                                                                    if (dataTable.Rows[y][column[0]].ToString() != "")
                                                                                    {
                                                                                        tblProprietary.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.BuyOrder = view;
                                                                                        }
                                                                                        else tblProprietary.BuyOrder = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.BuyVol = view;
                                                                                        }
                                                                                        else tblProprietary.BuyVol = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.SellOrder = view;
                                                                                        }
                                                                                        else tblProprietary.SellOrder = 0;
                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblProprietary.BuySellVol = view;
                                                                                        }
                                                                                        else tblProprietary.BuySellVol = 0;

                                                                                        tableproprietary5.Rows.Add(tblProprietary.TransDate, tblProprietary.CreateDate, tblProprietary.StockCode, tblProprietary.BuyOrder, tblProprietary.BuyVol, tblProprietary.SellOrder, tblProprietary.SellVol, tblProprietary.BuySellVol);
                                                                                    }
                                                                                }
                                                                                insert.InsertDB(tableproprietary5, config[configs.tableProprietary5]);

                                                                            }
                                                                        }

                                                                    }

                                                                }

                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else if (group7Value.Contains(configs.CK_GDLT) || group7Value.Contains(configs.TradingSummary) || group7Value.Contains(configs.Foreign3))
                                                    {
                                                        if (group6Value != "16.02.2009" && !group7Value.Contains(configs.Foreign3))
                                                        {
                                                            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                            {
                                                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                                {
                                                                    Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                    var dataSet = reader.AsDataSet();
                                                                    if (dataSet != null)
                                                                    {
                                                                        ConfigTable configTable = new ConfigTable();
                                                                        float view;
                                                                        string sheetName = config["Trading_Result:SheetName"];
                                                                        string[] sheetNames = sheetName.Split(',');
                                                                        foreach (string sheet in sheetNames)
                                                                        {
                                                                            if (dataSet.Tables.Contains(sheet))
                                                                            {
                                                                                if (sheet == "0" || sheet == "OM-CCQ (IFC'S)")
                                                                                {
                                                                                    DataSet dataSetX = null;
                                                                                    DataTable dataTable = null;
                                                                                    int beginRow = 0;
                                                                                    if (sheet == "0")
                                                                                    {
                                                                                        if (group6Date >= startDate14 && group6Date <= endDate14)
                                                                                        {
                                                                                            dataSetX = configTable.DatTenTrading_Result0(dataSet);
                                                                                            beginRow = Int32.Parse(config[configs.Trading_Result_0Row1]);
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            dataSetX = configTable.DatTenTrading_Result0(dataSet);
                                                                                            beginRow = Int32.Parse(config[configs.Trading_Result_0Row]);
                                                                                        }

                                                                                    }
                                                                                    else if (sheet == "OM-CCQ (IFC'S)")
                                                                                    {
                                                                                        dataSetX = configTable.DatTenTrading_Summary0(dataSet);
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_0Row1]);
                                                                                    }

                                                                                    string ColumOfDB = config[configs.Trading_Result_0Col];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tabletrading = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tabletrading.Columns.Add(column1);
                                                                                    }

                                                                                    dataTable = dataSetX.Tables[sheet.ToString()];

                                                                                    string beginRow1 = config[configs.Trading_Result_0Cell];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = beginRow; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableTrading_Result tbltrading = new ETableTrading_Result();
                                                                                        tbltrading.CreateDate = DateTime.Now;
                                                                                        tbltrading.TransDate = group6Date;
                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                        {
                                                                                            tbltrading.StockCode = dataTable.Rows[y][column[0]].ToString();
                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RefPrice = view;
                                                                                            }
                                                                                            else tbltrading.RefPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OpenPrice = view;
                                                                                            }
                                                                                            else tbltrading.OpenPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.HighPrice = view;
                                                                                            }
                                                                                            else tbltrading.HighPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.LowPrice = view;
                                                                                            }
                                                                                            else tbltrading.LowPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.AvgPrice = view;
                                                                                            }
                                                                                            else tbltrading.AvgPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.ClosePrice = view;
                                                                                            }
                                                                                            else tbltrading.ClosePrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.Change = view;
                                                                                            }
                                                                                            else tbltrading.Change = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TotalShareVol = view;
                                                                                            }
                                                                                            else tbltrading.TotalShareVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TotalValue = view;
                                                                                            }
                                                                                            else tbltrading.TotalValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.BidOrders = view;
                                                                                            }
                                                                                            else tbltrading.BidOrders = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OfferOrders = view;
                                                                                            }
                                                                                            else tbltrading.OfferOrders = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.BidVol = view;
                                                                                            }
                                                                                            else tbltrading.BidVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OfferVol = view;
                                                                                            }
                                                                                            else tbltrading.OfferVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RefPriceAfter = view;
                                                                                            }
                                                                                            else tbltrading.RefPriceAfter = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.CeilingPrice = view;
                                                                                            }
                                                                                            else tbltrading.CeilingPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.FloorPrice = view;
                                                                                            }
                                                                                            else tbltrading.FloorPrice = 0;

                                                                                            tabletrading.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.RefPrice, tbltrading.OpenPrice, tbltrading.HighPrice, tbltrading.LowPrice, tbltrading.AvgPrice, tbltrading.ClosePrice, tbltrading.Change, tbltrading.TotalShareVol, tbltrading.TotalValue, tbltrading.BidOrders, tbltrading.OfferOrders, tbltrading.BidVol, tbltrading.OfferVol, tbltrading.RefPriceAfter, tbltrading.CeilingPrice, tbltrading.FloorPrice);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tabletrading, config[configs.tableTrading]);

                                                                                }
                                                                                else if (sheet == "1" || sheet == "OM-CP(Stocks)")
                                                                                {
                                                                                    // var dataSetX = configTable.DatTenTrading_Result1(dataSet);
                                                                                    DataSet dataSetX = null;
                                                                                    DataTable dataTable = null;
                                                                                    int beginRow = 0;
                                                                                    if (sheet == "1")
                                                                                    {
                                                                                        dataSetX = configTable.DatTenTrading_Result1(dataSet);
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_1Row]);
                                                                                    }
                                                                                    else if (sheet == "OM-CP(Stocks)")
                                                                                    {
                                                                                        dataSetX = configTable.DatTenTrading_Summary1(dataSet);
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_1Row1]);
                                                                                    }

                                                                                    string ColumOfDB = config[configs.Trading_Result_1Col];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tabletrading1 = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tabletrading1.Columns.Add(column1);
                                                                                    }
                                                                                    dataTable = dataSetX.Tables[sheet.ToString()];

                                                                                    string beginRow1 = config[configs.Trading_Result_1Cell];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = beginRow; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableTrading_Result tbltrading = new ETableTrading_Result();
                                                                                        tbltrading.CreateDate = DateTime.Now;
                                                                                        tbltrading.TransDate = group6Date;
                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                        {
                                                                                            tbltrading.StockCode = dataTable.Rows[y][column[0]].ToString();
                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RefPrice = view;
                                                                                            }
                                                                                            else tbltrading.RefPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OpenPrice = view;
                                                                                            }
                                                                                            else tbltrading.OpenPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.HighPrice = view;
                                                                                            }
                                                                                            else tbltrading.HighPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.LowPrice = view;
                                                                                            }
                                                                                            else tbltrading.LowPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.AvgPrice = view;
                                                                                            }
                                                                                            else tbltrading.AvgPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.ClosePrice = view;
                                                                                            }
                                                                                            else tbltrading.ClosePrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.Change = view;
                                                                                            }
                                                                                            else tbltrading.Change = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TotalShareVol = view;
                                                                                            }
                                                                                            else tbltrading.TotalShareVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TotalValue = view;
                                                                                            }
                                                                                            else tbltrading.TotalValue = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.BidOrders = view;
                                                                                            }
                                                                                            else tbltrading.BidOrders = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OfferOrders = view;
                                                                                            }
                                                                                            else tbltrading.OfferOrders = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.BidVol = view;
                                                                                            }
                                                                                            else tbltrading.BidVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OfferVol = view;
                                                                                            }
                                                                                            else tbltrading.OfferVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RefPriceAfter = view;
                                                                                            }
                                                                                            else tbltrading.RefPriceAfter = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[15]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.CeilingPrice = view;
                                                                                            }
                                                                                            else tbltrading.CeilingPrice = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[16]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.FloorPrice = view;
                                                                                            }
                                                                                            else tbltrading.FloorPrice = 0;


                                                                                            tabletrading1.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.RefPrice, tbltrading.OpenPrice, tbltrading.HighPrice, tbltrading.LowPrice, tbltrading.AvgPrice, tbltrading.ClosePrice, tbltrading.Change, tbltrading.TotalShareVol, tbltrading.TotalValue, tbltrading.BidOrders, tbltrading.OfferOrders, tbltrading.BidVol, tbltrading.OfferVol, tbltrading.RefPriceAfter, tbltrading.CeilingPrice, tbltrading.FloorPrice);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tabletrading1, config[configs.tableTrading1]);

                                                                                }
                                                                                else if (sheet == "2" || sheet == "PT-TP(Bonds)")
                                                                                {
                                                                                    DataSet dataSetX = null;
                                                                                    DataTable dataTable = null;
                                                                                    int beginRow = 0;
                                                                                    if (sheet == "2")
                                                                                    {
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_2Row]);
                                                                                        dataSetX = configTable.DatTenTrading_Result2(dataSet);
                                                                                    }
                                                                                    else if (sheet == "PT-TP(Bonds)")
                                                                                    {
                                                                                        dataSetX = configTable.DatTenTrading_Summary2(dataSet);
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_2Row1]);
                                                                                    }

                                                                                    string ColumOfDB = config[configs.Trading_Result_2Col];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tabletrading2 = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tabletrading2.Columns.Add(column1);
                                                                                    }
                                                                                    dataTable = dataSetX.Tables[sheet.ToString()];

                                                                                    string beginRow1 = config[configs.Trading_Result_2Cell];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = beginRow; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableTrading_Result tbltrading = new ETableTrading_Result();

                                                                                        tbltrading.CreateDate = DateTime.Now;
                                                                                        tbltrading.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length < 10)
                                                                                        {
                                                                                            tbltrading.StockCode = (dataTable.Rows[y][column[0]]).ToString();


                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.OpenPrice = view;
                                                                                            }
                                                                                            else tbltrading.OpenPrice = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.ClosePrice = view;
                                                                                            }
                                                                                            else tbltrading.ClosePrice = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.CYield = view;
                                                                                            }
                                                                                            else tbltrading.CYield = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingVolume = view;
                                                                                            }
                                                                                            else tbltrading.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingValue = view;
                                                                                            }
                                                                                            else tbltrading.TradingValue = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.MaturityDate = view;
                                                                                            }
                                                                                            else tbltrading.MaturityDate = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.InterestRate = view;
                                                                                            }
                                                                                            else tbltrading.InterestRate = 0;
                                                                                            if (dataTable.Rows[y][column[8]].ToString() != "")
                                                                                            {
                                                                                                tbltrading.InterestPaymentMethod = dataTable.Rows[y][column[8]].ToString();
                                                                                            }
                                                                                            else tbltrading.InterestPaymentMethod = "NULL";
                                                                                            if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TimestoMaturity = view;
                                                                                            }
                                                                                            else tbltrading.TimestoMaturity = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.YTM = view;
                                                                                            }
                                                                                            else tbltrading.YTM = 0;

                                                                                            tabletrading2.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.OpenPrice, tbltrading.ClosePrice, tbltrading.CYield, tbltrading.TradingVolume, tbltrading.TradingValue, tbltrading.MaturityDate, tbltrading.InterestRate, tbltrading.InterestPaymentMethod, tbltrading.TimestoMaturity, tbltrading.YTM);

                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tabletrading2, config[configs.tableTrading2]);

                                                                                }
                                                                                else if (sheet == "3" || sheet == "PT-CP&CCQ(Stocks & IFCs)")
                                                                                {
                                                                                    DataSet dataSetX = null;
                                                                                    DataTable dataTable = null;
                                                                                    int beginRow = 0;
                                                                                    if (sheet == "3")
                                                                                    {
                                                                                        if (sheet == "3")
                                                                                        {
                                                                                            if (group6Date >= startDate14 && group6Date <= endDate14)
                                                                                            {
                                                                                                dataSetX = configTable.DatTenTrading_Result3(dataSet);
                                                                                                beginRow = Int32.Parse(config[configs.Trading_Result_3Row]);
                                                                                            }
                                                                                            else
                                                                                            {
                                                                                                dataSetX = configTable.DatTenTrading_Result3(dataSet);
                                                                                                beginRow = Int32.Parse(config[configs.Trading_Result_3Row1]);
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                    else if (sheet == "PT-CP&CCQ(Stocks & IFCs)")
                                                                                    {
                                                                                        dataSetX = configTable.DatTenTrading_Summary3(dataSet);
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_3Row]);
                                                                                    }
                                                                                    string ColumOfDB = config[configs.Trading_Result_3Col];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tabletrading3a = new DataTable();
                                                                                    DataTable tabletrading3b = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tabletrading3a.Columns.Add(column1);
                                                                                        tabletrading3b.Columns.Add(column1);
                                                                                    }
                                                                                    dataTable = dataSetX.Tables[sheet.ToString()];

                                                                                    string beginRow1 = config[configs.Trading_Result_3Cell];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = beginRow; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableTrading_Result tbltrading = new ETableTrading_Result();

                                                                                        tbltrading.CreateDate = DateTime.Now;
                                                                                        tbltrading.TransDate = group6Date;
                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length < 10)
                                                                                        {
                                                                                            tbltrading.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingVolume = view;
                                                                                            }
                                                                                            else tbltrading.TradingVolume = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingValue = view;
                                                                                            }
                                                                                            else tbltrading.TradingValue = 0;

                                                                                            tabletrading3a.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.TradingVolume, tbltrading.TradingValue);
                                                                                        }
                                                                                        if (dataTable.Rows[y][column[3]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0")
                                                                                        {
                                                                                            tbltrading.StockCode = (dataTable.Rows[y][column[3]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingVolume = view;
                                                                                            }
                                                                                            else tbltrading.TradingVolume = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingValue = view;
                                                                                            }
                                                                                            else tbltrading.TradingValue = 0;

                                                                                            tabletrading3b.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.TradingVolume, tbltrading.TradingValue);
                                                                                        }
                                                                                        else continue;
                                                                                        

                                                                                    }
                                                                                    StringBuilder sb = new StringBuilder();
                                                                                    sb.Append(configs.insertTbl).Append(" ").Append(config[configs.tableSession4]).Append(" ").Append(configs.valueTbl).Append(" ");

                                                                                    foreach (DataRow row in tabletrading3a.Rows)
                                                                                    {
                                                                                        sb.Append("(");
                                                                                        sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                                        sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                                        sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                                        sb.Append(row["TradingVolume"]).Append(",");
                                                                                        sb.Append(row["TradingValue"]);
                                                                                        sb.Append("),");
                                                                                    }

                                                                                    foreach (DataRow row in tabletrading3b.Rows)
                                                                                    {
                                                                                        sb.Append("(");
                                                                                        sb.Append("'").Append(row["TransDate"]).Append("',");
                                                                                        sb.Append("'").Append(row["CreateDate"]).Append("',");
                                                                                        sb.Append("'" + row["StockCode"] + "'").Append(",");
                                                                                        sb.Append(row["TradingVolume"]).Append(",");
                                                                                        sb.Append(row["TradingValue"]);
                                                                                        sb.Append("),");
                                                                                    }
                                                                                    if (tabletrading3a.Rows.Count > 0 || tabletrading3b.Rows.Count > 0)
                                                                                    {
                                                                                        command = new SqlCommand(sb.ToString().TrimEnd(','), sqlConnection);

                                                                                        command.ExecuteNonQuery();
                                                                                    }
                                                                                }
                                                                                else if (sheet == "4" || sheet == "CPQ (Treasury stocks)")
                                                                                {
                                                                                    DataSet dataSetX = null;
                                                                                    DataTable dataTable = null;
                                                                                    int beginRow = 0;
                                                                                    if (sheet == "4")
                                                                                    {
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_4Row]);
                                                                                        dataSetX = configTable.DatTenTrading_Result4(dataSet);
                                                                                    }
                                                                                    else if (sheet == "CPQ (Treasury stocks)")
                                                                                    {
                                                                                        dataSetX = configTable.DatTenTrading_Summary4(dataSet);
                                                                                        beginRow = Int32.Parse(config[configs.Trading_Result_4Row1]);
                                                                                    }
                                                                                    string ColumOfDB = config[configs.Trading_Result_4Col];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tabletrading4 = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tabletrading4.Columns.Add(column1);
                                                                                    }
                                                                                    dataTable = dataSetX.Tables[sheet.ToString()];

                                                                                    string beginRow1 = config[configs.Trading_Result_4Cell];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = beginRow; y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableTrading_Result tbltrading = new ETableTrading_Result();

                                                                                        tbltrading.CreateDate = DateTime.Now;
                                                                                        tbltrading.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length <= 10)
                                                                                        {
                                                                                            tbltrading.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                            if (dataTable.Rows[y][column[1]].ToString() != "")
                                                                                            {
                                                                                                tbltrading.BuySale = (dataTable.Rows[y][column[1]]).ToString();
                                                                                            }
                                                                                            else tbltrading.BuySale = "NULL";

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RegistrationVol = view;
                                                                                            }
                                                                                            else tbltrading.RegistrationVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.TradingVolume = view;
                                                                                            }
                                                                                            else tbltrading.TradingVolume = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RateRegistrationVol1 = view;
                                                                                            }
                                                                                            else tbltrading.RateRegistrationVol1 = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.AccumulatedTradingVol = view;
                                                                                            }
                                                                                            else tbltrading.AccumulatedTradingVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {

                                                                                                tbltrading.RateRegistrationVol2 = view;
                                                                                            }
                                                                                            else tbltrading.RateRegistrationVol2 = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RemainingVol = view;
                                                                                            }
                                                                                            else tbltrading.RemainingVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                            {
                                                                                                tbltrading.RateRegistrationVol3 = view;
                                                                                            }
                                                                                            else tbltrading.RateRegistrationVol3 = 0;
                                                                                            if (dataTable.Rows[y][column[9]].ToString() != "")
                                                                                            {
                                                                                                tbltrading.Deadline = dataTable.Rows[y][column[9]].ToString();
                                                                                            }
                                                                                            else tbltrading.Deadline = "NULL";

                                                                                            tabletrading4.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.BuySale, tbltrading.RegistrationVol, tbltrading.TradingVolume, tbltrading.RateRegistrationVol1, tbltrading.AccumulatedTradingVol, tbltrading.RateRegistrationVol2, tbltrading.RemainingVol, tbltrading.RateRegistrationVol3, tbltrading.Deadline);

                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tabletrading4, config[configs.tableTrading4]);
                                                                                }
                                                                                else if (sheet == "5")
                                                                                {
                                                                                    var dataSetX = configTable.DatTenForeign5(dataSet);
                                                                                    string ColumOfDB = config[configs.foreign5col];

                                                                                    string[] columns = ColumOfDB.Split(',');

                                                                                    DataTable tableforeign5 = new DataTable();
                                                                                    foreach (string column1 in columns)
                                                                                    {
                                                                                        tableforeign5.Columns.Add(column1);
                                                                                    }
                                                                                    DataTable dataTable = dataSetX.Tables["5"];

                                                                                    string beginRow1 = config[configs.foreign5cell];

                                                                                    string[] column = beginRow1.Split(',');

                                                                                    for (int y = Int32.Parse(config[configs.foreign5row]); y < dataTable.Rows.Count; y++)
                                                                                    {
                                                                                        ETableForeign tblForeign = new ETableForeign();
                                                                                        tblForeign.CreateDate = DateTime.Now;
                                                                                        tblForeign.TransDate = group6Date;

                                                                                        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString().Length <= 20)
                                                                                        {
                                                                                            tblForeign.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                            if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.TotalRoom = view;
                                                                                            }
                                                                                            else tblForeign.TotalRoom = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.CurrentRoom = view;
                                                                                            }
                                                                                            else tblForeign.CurrentRoom = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.ForeignOwnedRatio = view;
                                                                                            }
                                                                                            else tblForeign.ForeignOwnedRatio = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.StateOwnedRatio = view;
                                                                                            }
                                                                                            else tblForeign.StateOwnedRatio = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.OrderBuyPreOpen = view;
                                                                                            }
                                                                                            else tblForeign.OrderBuyPreOpen = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.OrderBuyCont = view;
                                                                                            }
                                                                                            else tblForeign.OrderBuyCont = 0;

                                                                                            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.OrderBuyPreClose = view;
                                                                                            }
                                                                                            else tblForeign.BuyVolPreClose = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.OrderSellPreOpen = view;
                                                                                            }
                                                                                            else tblForeign.OrderSellPreOpen = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[9]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.OpenSellCont = view;
                                                                                            }
                                                                                            else tblForeign.OpenSellCont = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[10]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.OpenSellClose = view;
                                                                                            }
                                                                                            else tblForeign.OpenSellClose = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[11]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.PutBuyVol = view;
                                                                                            }
                                                                                            else tblForeign.PutBuyVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[12]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.PutSellVol = view;
                                                                                            }
                                                                                            else tblForeign.PutSellVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[13]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.TotalBuyVol = view;
                                                                                            }
                                                                                            else tblForeign.TotalBuyVol = 0;
                                                                                            if (float.TryParse(dataTable.Rows[y][column[14]].ToString(), out view))
                                                                                            {
                                                                                                tblForeign.TotalSellVol = view;
                                                                                            }
                                                                                            else tblForeign.TotalSellVol = 0;

                                                                                            tableforeign5.Rows.Add(tblForeign.CreateDate, tblForeign.TransDate, tblForeign.StockCode, tblForeign.TotalRoom, tblForeign.CurrentRoom, tblForeign.ForeignOwnedRatio, tblForeign.StateOwnedRatio, tblForeign.OrderBuyPreOpen, tblForeign.OrderBuyCont, tblForeign.OrderBuyPreClose, tblForeign.OrderSellPreOpen, tblForeign.OpenSellCont, tblForeign.OpenSellClose, tblForeign.PutBuyVol, tblForeign.PutSellVol, tblForeign.TotalBuyVol, tblForeign.TotalSellVol);
                                                                                        }
                                                                                    }
                                                                                    insert.InsertDB(tableforeign5, config[configs.tableForeign5]);
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        throw new Exception(configs.ERROR);
                                                                    }
                                                                }
                                                            }
                                                            break;
                                                        }

                                                    }
                                                    else if (group7Value.Contains(configs.PutThrough) || group7Value.Contains(configs.PT))
                                                    {
                                                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                                                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                                                        {
                                                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                                                            {
                                                                Console.WriteLine("Đang đọc file Excel: " + filePath + "\n");

                                                                var dataSet = reader.AsDataSet();
                                                                if (dataSet != null)
                                                                {
                                                                    ConfigTable configTable = new ConfigTable();
                                                                    float view;
                                                                    string sheetName = config[configs.PutThough_TreasurySheetName];
                                                                    string[] sheetNames = sheetName.Split(',');
                                                                    foreach (string sheet in sheetNames)
                                                                    {
                                                                        if (dataSet.Tables.Contains(sheet))
                                                                        {
                                                                            if (sheet == "PT")
                                                                            {
                                                                                var dataSetX = configTable.DatTenPutThoughPT(dataSet);
                                                                               // string ColumOfDB1 = config[configs.PTcol1];
                                                                                string ColumOfDB2 = config[configs.PTcol2];
                                                                              //  string[] columns1 = ColumOfDB1.Split(',');
                                                                                string[] columns2 = ColumOfDB2.Split(',');
                                                                              //  DataTable tabletrading3a = new DataTable();
                                                                                DataTable tabletrading3b = new DataTable();
                                                                                //foreach (string column1 in columns1)
                                                                                //{
                                                                                //    tabletrading3a.Columns.Add(column1);
                                                                                //}
                                                                                foreach (string column1 in columns2)
                                                                                {
                                                                                    tabletrading3b.Columns.Add(column1);
                                                                                }
                                                                                DataTable dataTable = dataSetX.Tables["PT"];

                                                                                string beginRow1 = config[configs.PTCell];

                                                                                string[] column = beginRow1.Split(',');

                                                                                //for (int y = Int32.Parse(config[configs.PTRow]); y < dataTable.Rows.Count; y++)
                                                                                //{
                                                                                //    ETablePutThough_Treasury tblput = new ETablePutThough_Treasury();

                                                                                //    tblput.CreateDate = DateTime.Now;
                                                                                //    tblput.TransDate = group6Date;
                                                                                //    if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length <= 15)
                                                                                //    {
                                                                                //        tblput.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                                //        if (float.TryParse(dataTable.Rows[y][column[1]].ToString(), out view))
                                                                                //        {
                                                                                //            tblput.TradingVolume = view;
                                                                                //        }
                                                                                //        else tblput.TradingVolume = 0;

                                                                                //        if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                                //        {
                                                                                //            tblput.TradingValue = view;
                                                                                //        }
                                                                                //        else tblput.TradingValue = 0;

                                                                                //        tabletrading3a.Rows.Add(tblput.TransDate, tblput.CreateDate, tblput.StockCode, tblput.TradingVolume, tblput.TradingValue);
                                                                                //    }
                                                                                //}

                                                                                for (int y = Int32.Parse(config[configs.PTRow]); y < dataTable.Rows.Count; y++)
                                                                                {
                                                                                    ETablePutThough_Treasury tblput = new ETablePutThough_Treasury();

                                                                                    tblput.CreateDate = DateTime.Now;
                                                                                    tblput.TransDate = group6Date;
                                                                                    if (y == 8)
                                                                                    {
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblput.TotalTrading = view;
                                                                                        }
                                                                                        else tblput.TotalTrading = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblput.RateTotalTrading = view;
                                                                                        }
                                                                                        else tblput.RateTotalTrading = 0;

                                                                                        tblput.Type = "TotalVolumeEntireMarket";
                                                                                    }
                                                                                    else if (y == 9)
                                                                                    {
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblput.TotalTrading = view;
                                                                                        }
                                                                                        else tblput.TotalTrading = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblput.RateTotalTrading = view;
                                                                                        }
                                                                                        else tblput.RateTotalTrading = 0;

                                                                                        tblput.Type = "TotalVolumeOrderMatching";
                                                                                    }
                                                                                    else if (y == 10)
                                                                                    {
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblput.TotalTrading = view;
                                                                                        }
                                                                                        else tblput.TotalTrading = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblput.RateTotalTrading = view;
                                                                                        }
                                                                                        else tblput.RateTotalTrading = 0;

                                                                                        tblput.Type = "TotalVolumePutThough";
                                                                                    }
                                                                                    else if (y == 13)
                                                                                    {
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblput.TotalTrading = view;
                                                                                        }
                                                                                        else tblput.TotalTrading = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblput.RateTotalTrading = view;
                                                                                        }
                                                                                        else tblput.RateTotalTrading = 0;

                                                                                        tblput.Type = "TotalValueEntireMarket";
                                                                                    }
                                                                                    else if (y == 14)
                                                                                    {
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblput.TotalTrading = view;
                                                                                        }
                                                                                        else tblput.TotalTrading = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblput.RateTotalTrading = view;
                                                                                        }
                                                                                        else tblput.RateTotalTrading = 0;

                                                                                        tblput.Type = "TotalValueOrderMatching";
                                                                                    }
                                                                                    else if (y == 15)
                                                                                    {
                                                                                        if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                                        {
                                                                                            tblput.TotalTrading = view;
                                                                                        }
                                                                                        else tblput.TotalTrading = 0;

                                                                                        if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                                        {
                                                                                            tblput.RateTotalTrading = view;
                                                                                        }
                                                                                        else tblput.RateTotalTrading = 0;

                                                                                        tblput.Type = "TotalValuePutThough";
                                                                                    }
                                                                                    else continue;

                                                                                    tabletrading3b.Rows.Add(tblput.TransDate, tblput.CreateDate, tblput.TotalTrading, tblput.RateTotalTrading, tblput.Type);

                                                                                }

                                                                             //   insert.InsertDB(tabletrading3a, config[configs.tableSession4]);
                                                                                insert.InsertDB(tabletrading3b, config[configs.tableTotalTrading]);

                                                                            }
                                                                            //else if (sheet == "MBL")
                                                                            //{
                                                                            //    var dataSetX = configTable.DatTenPutThoughMBL(dataSet);
                                                                            //    string ColumOfDB = config[configs.MBLcol];

                                                                            //    string[] columns = ColumOfDB.Split(',');

                                                                            //    DataTable tabletrading4 = new DataTable();
                                                                            //    foreach (string column1 in columns)
                                                                            //    {
                                                                            //        tabletrading4.Columns.Add(column1);
                                                                            //    }
                                                                            //    DataTable dataTable = dataSetX.Tables["MBL"];

                                                                            //    string beginRow1 = config[configs.MBLCell];

                                                                            //    string[] column = beginRow1.Split(',');

                                                                            //    for (int y = Int32.Parse(config[configs.MBLRow]); y < dataTable.Rows.Count; y++)
                                                                            //    {
                                                                            //        ETableTrading_Result tbltrading = new ETableTrading_Result();

                                                                            //        tbltrading.CreateDate = DateTime.Now;
                                                                            //        tbltrading.TransDate = group6Date;

                                                                            //        if (dataTable.Rows[y][column[0]].ToString() != "" && dataTable.Rows[y][column[0]].ToString() != "0" && dataTable.Rows[y][column[0]].ToString().Length < 15)
                                                                            //        {
                                                                            //            tbltrading.StockCode = (dataTable.Rows[y][column[0]]).ToString();

                                                                            //            if (dataTable.Rows[y][column[1]].ToString() != "")
                                                                            //            {
                                                                            //                tbltrading.BuySale = (dataTable.Rows[y][column[1]]).ToString();
                                                                            //            }
                                                                            //            else tbltrading.BuySale = "NULL";

                                                                            //            if (float.TryParse(dataTable.Rows[y][column[2]].ToString(), out view))
                                                                            //            {
                                                                            //                tbltrading.RegistrationVol = view;
                                                                            //            }
                                                                            //            else tbltrading.RegistrationVol = 0;
                                                                            //            if (float.TryParse(dataTable.Rows[y][column[3]].ToString(), out view))
                                                                            //            {
                                                                            //                tbltrading.TradingVolume = view;
                                                                            //            }
                                                                            //            else tbltrading.TradingVolume = 0;
                                                                            //            if (float.TryParse(dataTable.Rows[y][column[4]].ToString(), out view))
                                                                            //            {
                                                                            //                tbltrading.RateRegistrationVol1 = view;
                                                                            //            }
                                                                            //            else tbltrading.RateRegistrationVol1 = 0;
                                                                            //            if (float.TryParse(dataTable.Rows[y][column[5]].ToString(), out view))
                                                                            //            {
                                                                            //                tbltrading.AccumulatedTradingVol = view;
                                                                            //            }
                                                                            //            else tbltrading.AccumulatedTradingVol = 0;
                                                                            //            if (float.TryParse(dataTable.Rows[y][column[6]].ToString(), out view))
                                                                            //            {

                                                                            //                tbltrading.RateRegistrationVol2 = view;
                                                                            //            }
                                                                            //            else tbltrading.RateRegistrationVol2 = 0;
                                                                            //            if (float.TryParse(dataTable.Rows[y][column[7]].ToString(), out view))
                                                                            //            {
                                                                            //                tbltrading.RemainingVol = view;
                                                                            //            }
                                                                            //            else tbltrading.RemainingVol = 0;
                                                                            //            if (float.TryParse(dataTable.Rows[y][column[8]].ToString(), out view))
                                                                            //            {
                                                                            //                tbltrading.RateRegistrationVol3 = view;
                                                                            //            }
                                                                            //            else tbltrading.RateRegistrationVol3 = 0;
                                                                            //            if (dataTable.Rows[y][column[9]].ToString() != "")
                                                                            //            {
                                                                            //                tbltrading.Deadline = dataTable.Rows[y][column[9]].ToString();
                                                                            //            }
                                                                            //            else tbltrading.Deadline = "NULL";

                                                                            //            tabletrading4.Rows.Add(tbltrading.TransDate, tbltrading.CreateDate, tbltrading.StockCode, tbltrading.BuySale, tbltrading.RegistrationVol, tbltrading.TradingVolume, tbltrading.RateRegistrationVol1, tbltrading.AccumulatedTradingVol, tbltrading.RateRegistrationVol2, tbltrading.RemainingVol, tbltrading.RateRegistrationVol3, tbltrading.Deadline);

                                                                            //        }
                                                                            //    }
                                                                            //    insert.InsertDB(tabletrading4, config[configs.tableTrading4]);

                                                                            //}
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception(configs.ERROR);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        // ngày khác
                                                        //Console.WriteLine("file khác " + filePath);
                                                        //break;
                                                    }
                                                }

                                                else
                                                {
                                                    Console.WriteLine("Invalid start or end date.");
                                                    break;
                                                }

                                            }

                                        }
                                    }

                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            writer.WriteLine($"Error reading file {filePath}: {ex.Message} intime: {DateTime.Now}");
                            Console.WriteLine($"Error reading file {filePath}: {ex.Message}");
                        }

                    }

                }
            }
            catch (Exception ex)
            {
            }

            //  }

        }

    }
}
