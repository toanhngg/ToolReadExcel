using System.Data;

namespace ApiExcelToDB.Entities
{
    public class ConfigTable : ConfigApp
    {
        public DataSet DatTen1(DataSet dataSet)
        {
            // sheet 1
            dataSet.Tables["1"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["1"].Columns[3].ColumnName = "D8"; // PriorDayClose
            dataSet.Tables["1"].Columns[4].ColumnName = "E8"; // SessionClose
            dataSet.Tables["1"].Columns[7].ColumnName = "H8";   // Change
            dataSet.Tables["1"].Columns[8].ColumnName = "I8";   //TradingVolume
            dataSet.Tables["1"].Columns[9].ColumnName = "J8";  //TradingValue
            return dataSet;
        }

        public DataSet DatTen1CW(DataSet dataSet)
        {
            dataSet.Tables["1.CW"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["1.CW"].Columns[3].ColumnName = "D8"; // PriorDayClose
            dataSet.Tables["1.CW"].Columns[4].ColumnName = "E8"; // SessionClose
            dataSet.Tables["1.CW"].Columns[7].ColumnName = "H8";   // Change
            dataSet.Tables["1.CW"].Columns[8].ColumnName = "I8";   //TradingVolume
            dataSet.Tables["1.CW"].Columns[9].ColumnName = "J8";  //TradingValue
            return dataSet;
        }
        public DataSet DatTen2(DataSet dataSet)
        {
            dataSet.Tables["2"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["2"].Columns[3].ColumnName = "D8"; //PriorDayClose
            dataSet.Tables["2"].Columns[4].ColumnName = "E8"; //SessionHigh
            dataSet.Tables["2"].Columns[5].ColumnName = "F8"; //SessionAverage
            dataSet.Tables["2"].Columns[6].ColumnName = "G8"; //SessionLow
            dataSet.Tables["2"].Columns[7].ColumnName = "H8";// SessionClose
            dataSet.Tables["2"].Columns[8].ColumnName = "I8"; // TradingVolume
            dataSet.Tables["2"].Columns[9].ColumnName = "J8"; // TradingValue

            return dataSet;
        }
        public DataSet DatTen2CW(DataSet dataSet)
        {
            dataSet.Tables["2.CW"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["2.CW"].Columns[3].ColumnName = "D8"; //PriorDayClose
            dataSet.Tables["2.CW"].Columns[4].ColumnName = "E8"; //SessionHigh
            dataSet.Tables["2.CW"].Columns[5].ColumnName = "F8"; //SessionAverage
            dataSet.Tables["2.CW"].Columns[6].ColumnName = "G8"; //SessionLow
            dataSet.Tables["2.CW"].Columns[7].ColumnName = "H8";// SessionClose
            dataSet.Tables["2.CW"].Columns[8].ColumnName = "I8"; // TradingVolume
            dataSet.Tables["2.CW"].Columns[9].ColumnName = "J8"; // TradingValue

            return dataSet;
        }
        public DataSet DatTen2ODD_OM(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["2.ODD_OM"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["2.ODD_OM"].Columns[2].ColumnName = "C8"; //TradingVolume
            dataSet.Tables["2.ODD_OM"].Columns[3].ColumnName = "D8"; //TradingValue 

            return dataSet;
        }
        public DataSet DatTen3(DataSet dataSet)
        {
            // sheet 1
            dataSet.Tables["3"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["3"].Columns[3].ColumnName = "D8"; //PriorDayClose
            dataSet.Tables["3"].Columns[4].ColumnName = "E8"; //TodayClose 
            dataSet.Tables["3"].Columns[7].ColumnName = "H8"; //Change
            dataSet.Tables["3"].Columns[8].ColumnName = "I8"; //TradingVolume
            dataSet.Tables["3"].Columns[9].ColumnName = "J8"; //TradingValue
            dataSet.Tables["3"].Columns[10].ColumnName = "K8"; //Totalvolume
            dataSet.Tables["3"].Columns[11].ColumnName = "L8"; //Totalvalue
            dataSet.Tables["3"].Columns[12].ColumnName = "M8"; //ListedShares
            dataSet.Tables["3"].Columns[13].ColumnName = "N8"; //OutstandingShares
            dataSet.Tables["3"].Columns[14].ColumnName = "O8"; //AdjustedOutstandingShares
            dataSet.Tables["3"].Columns[15].ColumnName = "P8"; //Marketcap

            return dataSet;
        }
        public DataSet DatTen3CW(DataSet dataSet)
        {
            // sheet 1
            dataSet.Tables["3.CW"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["3.CW"].Columns[3].ColumnName = "D8"; //PriorDayClose
            dataSet.Tables["3.CW"].Columns[4].ColumnName = "E8"; //TodayClose 
            dataSet.Tables["3.CW"].Columns[7].ColumnName = "H8"; //Change
            dataSet.Tables["3.CW"].Columns[8].ColumnName = "I8"; //TradingVolume
            dataSet.Tables["3.CW"].Columns[9].ColumnName = "J8"; //TradingValue
            dataSet.Tables["3.CW"].Columns[10].ColumnName = "K8"; //Totalvolume
            dataSet.Tables["3.CW"].Columns[11].ColumnName = "L8"; //Totalvalue
            dataSet.Tables["3.CW"].Columns[12].ColumnName = "M8"; //ListedShares


            return dataSet;
        }
        public DataSet DatTen4(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["4"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["4"].Columns[2].ColumnName = "C8"; //TradingVolume
            dataSet.Tables["4"].Columns[3].ColumnName = "D8"; //TradingValue


            return dataSet;
        }
        public DataSet DatTen4CW(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["4.CW"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["4.CW"].Columns[2].ColumnName = "C8"; //TradingVolume
            dataSet.Tables["4.CW"].Columns[3].ColumnName = "D8"; //TradingValue


            return dataSet;
        }
        public DataSet DatTenODD_PT(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["4.ODD_PT"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["4.ODD_PT"].Columns[2].ColumnName = "C8"; //TradingVolume
            dataSet.Tables["4.ODD_PT"].Columns[3].ColumnName = "D8"; //TradingValue


            return dataSet;
        }
        public DataSet DatTenHOSEINDEX(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["HOSEINDEX"].Columns[1].ColumnName = "B8"; //IndexName
            dataSet.Tables["HOSEINDEX"].Columns[2].ColumnName = "C8"; //OpenIndexValue
            dataSet.Tables["HOSEINDEX"].Columns[3].ColumnName = "D8"; //CloseIndexValue
            dataSet.Tables["HOSEINDEX"].Columns[4].ColumnName = "E8"; //High
            dataSet.Tables["HOSEINDEX"].Columns[5].ColumnName = "F8"; //Low
            dataSet.Tables["HOSEINDEX"].Columns[6].ColumnName = "G8"; //UpDown
            dataSet.Tables["HOSEINDEX"].Columns[7].ColumnName = "H8"; //Change
            dataSet.Tables["HOSEINDEX"].Columns[8].ColumnName = "I8"; //TradingVolume
            dataSet.Tables["HOSEINDEX"].Columns[9].ColumnName = "J8"; //TradingValue
            dataSet.Tables["HOSEINDEX"].Columns[10].ColumnName = "K8"; //Marketcap



            return dataSet;
        }
        //// bảng basic idititors

        public DataSet DatTenBasic1(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["1"].Columns[1].ColumnName = "B13"; //StockCode
            dataSet.Tables["1"].Columns[2].ColumnName = "C13"; //PriorDayClose
            dataSet.Tables["1"].Columns[3].ColumnName = "D13"; //52wRecordHigh
            dataSet.Tables["1"].Columns[4].ColumnName = "E13"; //52wRecordLow
            dataSet.Tables["1"].Columns[5].ColumnName = "F13"; //AverageOutstandingShares
            dataSet.Tables["1"].Columns[6].ColumnName = "G13"; //PrimaryEPS
            dataSet.Tables["1"].Columns[7].ColumnName = "H13"; //Notes
            dataSet.Tables["1"].Columns[12].ColumnName = "M13"; //AdjustedEPS
            dataSet.Tables["1"].Columns[14].ColumnName = "O13"; //P/E
            dataSet.Tables["1"].Columns[15].ColumnName = "P13"; //Dividend
            dataSet.Tables["1"].Columns[16].ColumnName = "Q13"; //Dividend/MarketPrice
            dataSet.Tables["1"].Columns[17].ColumnName = "R13"; //ReturnOnTotalAssets
            dataSet.Tables["1"].Columns[18].ColumnName = "S13"; //ReturnOnEquity
            dataSet.Tables["1"].Columns[19].ColumnName = "T13"; //ListedShares
            dataSet.Tables["1"].Columns[20].ColumnName = "U13"; //OutstandingShares
            dataSet.Tables["1"].Columns[21].ColumnName = "V13"; //ChangeOutstandingShares
            dataSet.Tables["1"].Columns[22].ColumnName = "W13"; //AdjustedOutstandingShares
            dataSet.Tables["1"].Columns[23].ColumnName = "X13"; //TurnoverRatio

            return dataSet;
        }
        public DataSet DatTenBasicOld(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["1"].Columns[1].ColumnName = "B9"; //StockCode
            dataSet.Tables["1"].Columns[2].ColumnName = "C9"; //52wRecordHigh
            dataSet.Tables["1"].Columns[3].ColumnName = "D9"; //52wRecordLow
            dataSet.Tables["1"].Columns[4].ColumnName = "E9"; //AverageOutstandingShares
            dataSet.Tables["1"].Columns[5].ColumnName = "F9"; //PrimaryEPS
            dataSet.Tables["1"].Columns[6].ColumnName = "G9"; //Notes
            dataSet.Tables["1"].Columns[11].ColumnName = "L9"; //AdjustedEPS
            dataSet.Tables["1"].Columns[13].ColumnName = "N9"; //P/E
            dataSet.Tables["1"].Columns[14].ColumnName = "O9"; //Dividend
            dataSet.Tables["1"].Columns[15].ColumnName = "P9"; //Dividend/MarketPrice
            dataSet.Tables["1"].Columns[16].ColumnName = "Q9"; //ListedShares
            dataSet.Tables["1"].Columns[17].ColumnName = "R9"; //OutstandingShares
            dataSet.Tables["1"].Columns[18].ColumnName = "S9"; //ChangeOutstandingShares
            dataSet.Tables["1"].Columns[19].ColumnName = "T9"; //AdjustedOutstandingShares
            dataSet.Tables["1"].Columns[21].ColumnName = "V9"; //MtkCap

            return dataSet;
        }
        public DataSet DatTenBasicOld2(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["1"].Columns[1].ColumnName = "B9"; //StockCode
            dataSet.Tables["1"].Columns[2].ColumnName = "C9"; //52wRecordHigh
            dataSet.Tables["1"].Columns[3].ColumnName = "D9"; //52wRecordLow
            dataSet.Tables["1"].Columns[4].ColumnName = "E9"; //AverageOutstandingShares
            dataSet.Tables["1"].Columns[5].ColumnName = "F9"; //PrimaryEPS
            dataSet.Tables["1"].Columns[6].ColumnName = "G9"; //Notes
            dataSet.Tables["1"].Columns[11].ColumnName = "L9"; //AdjustedEPS
            dataSet.Tables["1"].Columns[13].ColumnName = "N9"; //P/E
            dataSet.Tables["1"].Columns[14].ColumnName = "O9"; //Dividend
            dataSet.Tables["1"].Columns[15].ColumnName = "P9"; //Dividend/MarketPrice
            dataSet.Tables["1"].Columns[16].ColumnName = "Q9"; //ListedShares
            dataSet.Tables["1"].Columns[17].ColumnName = "R9"; //OutstandingShares
            dataSet.Tables["1"].Columns[18].ColumnName = "S9"; //AdjustedOutstandingShares
            dataSet.Tables["1"].Columns[20].ColumnName = "U9"; //MtkCap

            return dataSet;
        }
        public DataSet DatTenBasicOld1(DataSet dataSet)
        {
            // sheet 1
            //A8,B8,C8,D8
            dataSet.Tables["1"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["1"].Columns[2].ColumnName = "C6"; //AverageOutstandingShares
            dataSet.Tables["1"].Columns[3].ColumnName = "D6"; //PrimaryEPS
            dataSet.Tables["1"].Columns[4].ColumnName = "E6"; //Notes
            dataSet.Tables["1"].Columns[9].ColumnName = "J6"; //AdjustedEPS
            dataSet.Tables["1"].Columns[11].ColumnName = "L6"; //PE
            dataSet.Tables["1"].Columns[12].ColumnName = "M6"; //Dividend
            dataSet.Tables["1"].Columns[13].ColumnName = "N6"; //Dividend/MarketPrice
            dataSet.Tables["1"].Columns[14].ColumnName = "O6"; //ListedShares
            dataSet.Tables["1"].Columns[15].ColumnName = "P6"; //OutstandingShares
            dataSet.Tables["1"].Columns[17].ColumnName = "R6"; //MktCap

            return dataSet;
        }
        public DataSet DatTenBasicOld3(DataSet dataSet)
        {
            // bat dau tu 9 b- t

            dataSet.Tables["1"].Columns[1].ColumnName = "B9"; //StockCode
            dataSet.Tables["1"].Columns[2].ColumnName = "C9"; //AverageOutstandingShares
            dataSet.Tables["1"].Columns[3].ColumnName = "D9"; //PrimaryEPS
            dataSet.Tables["1"].Columns[4].ColumnName = "E9"; //Notes
            dataSet.Tables["1"].Columns[5].ColumnName = "F9"; //AdjustedEPS
            dataSet.Tables["1"].Columns[6].ColumnName = "G9"; //PE
            dataSet.Tables["1"].Columns[11].ColumnName = "L9"; //Dividend
            dataSet.Tables["1"].Columns[13].ColumnName = "N9"; //Dividend/MarketPrice
            dataSet.Tables["1"].Columns[14].ColumnName = "O9"; //ListedShares
            dataSet.Tables["1"].Columns[15].ColumnName = "P9"; //OutstandingShares
            dataSet.Tables["1"].Columns[16].ColumnName = "Q9"; //OutstandingShares
            dataSet.Tables["1"].Columns[17].ColumnName = "R9"; //MktCap
            dataSet.Tables["1"].Columns[19].ColumnName = "T9"; //MktCap


            return dataSet;
        }
        //public DataSet DatTenBasicOld4(DataSet dataSet)
        //{// bat dau tu 9 b- t

        //    dataSet.Tables["1"].Columns[1].ColumnName = "B6"; //StockCode
        //    dataSet.Tables["1"].Columns[2].ColumnName = "C6"; //AverageOutstandingShares
        //    dataSet.Tables["1"].Columns[3].ColumnName = "D6"; //PrimaryEPS
        //    dataSet.Tables["1"].Columns[4].ColumnName = "E6"; //Notes
        //    dataSet.Tables["1"].Columns[9].ColumnName = "J6"; //AdjustedEPS
        //    dataSet.Tables["1"].Columns[11].ColumnName = "L6"; //PE
        //    dataSet.Tables["1"].Columns[12].ColumnName = "M6"; //Dividend
        //    dataSet.Tables["1"].Columns[13].ColumnName = "N6"; //Dividend/MarketPrice
        //    dataSet.Tables["1"].Columns[14].ColumnName = "O6"; //ListedShares
        //    dataSet.Tables["1"].Columns[15].ColumnName = "P6"; //OutstandingShares
        //    dataSet.Tables["1"].Columns[17].ColumnName = "R6"; //MktCap


        //    return dataSet;
        //}
        //TransDate,CreateDate,StockCode,wRecordHigh,wRecordLow,AverageOutstandingShares,PrimaryEPS,Notes,AdjustedEPS,PE,Dividend,DividendMarketPrice,ListedShares,OutstandingShares,ChangeOutstandingShares,AdjustedOutstandingShares,MtkCap
        public DataSet DatTenOrderChitiet(DataSet dataSet)
        {
            // sheet 1
            dataSet.Tables["Chi tiet (Details)"].Columns[0].ColumnName = "A8"; //StockCode
            dataSet.Tables["Chi tiet (Details)"].Columns[1].ColumnName = "B8"; //BuyingOrders
            dataSet.Tables["Chi tiet (Details)"].Columns[2].ColumnName = "C8"; //BuyingVolume
            dataSet.Tables["Chi tiet (Details)"].Columns[3].ColumnName = "D8"; //SellingOrders
            dataSet.Tables["Chi tiet (Details)"].Columns[4].ColumnName = "E8"; //SellingVolume
            dataSet.Tables["Chi tiet (Details)"].Columns[5].ColumnName = "F8"; //TradingVolume
            dataSet.Tables["Chi tiet (Details)"].Columns[6].ColumnName = "G8"; //BuySellVolume


            return dataSet;
        }
        public DataSet DatTenOrderCW(DataSet dataSet)
        {
            // sheet 1
            dataSet.Tables["CW"].Columns[0].ColumnName = "A8"; //StockCode
            dataSet.Tables["CW"].Columns[1].ColumnName = "B8"; //BuyingOrders
            dataSet.Tables["CW"].Columns[2].ColumnName = "C8"; //BuyingVolume
            dataSet.Tables["CW"].Columns[3].ColumnName = "D8"; //SellingOrders
            dataSet.Tables["CW"].Columns[4].ColumnName = "E8"; //SellingVolume
            dataSet.Tables["CW"].Columns[5].ColumnName = "F8"; //TradingVolume
            dataSet.Tables["CW"].Columns[6].ColumnName = "G8"; //BuySellVolume


            return dataSet;
        }
        public DataSet DatTenOrderODD(DataSet dataSet)
        {
            // sheet 1
            dataSet.Tables["ODD"].Columns[0].ColumnName = "A8"; //StockCode
            dataSet.Tables["ODD"].Columns[1].ColumnName = "B8"; //BuyingOrders
            dataSet.Tables["ODD"].Columns[2].ColumnName = "C8"; //BuyingVolume
            dataSet.Tables["ODD"].Columns[3].ColumnName = "D8"; //SellingOrders
            dataSet.Tables["ODD"].Columns[4].ColumnName = "E8"; //SellingVolume
            dataSet.Tables["ODD"].Columns[5].ColumnName = "F8"; //TradingVolume
            dataSet.Tables["ODD"].Columns[6].ColumnName = "G8"; //BuySellVolume


            return dataSet;
        }
        //TKGD NDTNN (Foreign Trading)
        public DataSet DatTenForeign1(DataSet dataSet)
        {
            // Type,Tradingvolume,rateEntireMaket,TradingValue,ratEntireMaket2
            dataSet.Tables["1"].Columns[1].ColumnName = "B11"; //Tradingvolume
            dataSet.Tables["1"].Columns[2].ColumnName = "C11"; //rateEntireMaket
            dataSet.Tables["1"].Columns[3].ColumnName = "D11"; //TradingValue
            dataSet.Tables["1"].Columns[4].ColumnName = "E11"; //ratEntireMaket2



            return dataSet;
        }
        public DataSet DatTenForeign2(DataSet dataSet)
        {
            // STT,StockCode,TotalRoom,CurrentRoom,ForeignOwnedRatio,StateOwnedRadio,BuyVolPreOpen,BuyVolCount,BuyVolPreClose,BuyValue,SellVolPreOpen,SellVolCount,SellVolPreClose,SellValue,PutBuyVol,PutBuyVal,PutSellVol,PutSellVal
            dataSet.Tables["2"].Columns[1].ColumnName = "B11"; //StockCode
            dataSet.Tables["2"].Columns[2].ColumnName = "C11"; //TotalRoom
            dataSet.Tables["2"].Columns[3].ColumnName = "D11"; //CurrentRoom
            dataSet.Tables["2"].Columns[4].ColumnName = "E11"; //ForeignOwnedRatio
            dataSet.Tables["2"].Columns[5].ColumnName = "F11"; //StateOwnedRadio
            dataSet.Tables["2"].Columns[6].ColumnName = "G11"; //BuyVolPreOpen
            dataSet.Tables["2"].Columns[7].ColumnName = "H11"; //BuyVolCount
            dataSet.Tables["2"].Columns[8].ColumnName = "I11"; //BuyVolPreClose
            dataSet.Tables["2"].Columns[9].ColumnName = "J11"; //BuyValue
            dataSet.Tables["2"].Columns[10].ColumnName = "K11"; //SellVolPreOpen
            dataSet.Tables["2"].Columns[11].ColumnName = "L11"; //SellVolCount
            dataSet.Tables["2"].Columns[12].ColumnName = "M11"; //SellVolPreClose
            dataSet.Tables["2"].Columns[13].ColumnName = "N11"; //SellValue,,,PutSellVol,PutSellVal
            dataSet.Tables["2"].Columns[14].ColumnName = "O11"; //PutBuyVol
            dataSet.Tables["2"].Columns[15].ColumnName = "P11"; //PutBuyVal
            dataSet.Tables["2"].Columns[16].ColumnName = "Q11"; //PutSellVol
            dataSet.Tables["2"].Columns[17].ColumnName = "R11"; //PutSellVal


            return dataSet;
        }
        public DataSet DatTenForeign2cu(DataSet dataSet)
        {
            // STT,StockCode,TotalRoom,CurrentRoom,ForeignOwnedRatio,StateOwnedRadio,BuyVolPreOpen,BuyVolCount,BuyVolPreClose,BuyValue,SellVolPreOpen,SellVolCount,SellVolPreClose,SellValue,PutBuyVol,PutBuyVal,PutSellVol,PutSellVal
            dataSet.Tables["2"].Columns[0].ColumnName = "A11"; //StockCode
            dataSet.Tables["2"].Columns[1].ColumnName = "B11"; //TotalRoom
            dataSet.Tables["2"].Columns[2].ColumnName = "C11"; //CurrentRoom
            dataSet.Tables["2"].Columns[3].ColumnName = "D11"; //ForeignOwnedRatio
            dataSet.Tables["2"].Columns[4].ColumnName = "E11"; //StateOwnedRadio
            dataSet.Tables["2"].Columns[5].ColumnName = "F11"; //BuyVolPreOpen
            dataSet.Tables["2"].Columns[6].ColumnName = "G11"; //BuyVolCount
            dataSet.Tables["2"].Columns[7].ColumnName = "H11"; //BuyVolPreClose
            dataSet.Tables["2"].Columns[8].ColumnName = "I11"; //BuyValue
            dataSet.Tables["2"].Columns[9].ColumnName = "J11"; //SellVolPreOpen
            dataSet.Tables["2"].Columns[10].ColumnName = "K11"; //SellVolCount
            dataSet.Tables["2"].Columns[11].ColumnName = "L11"; //SellVolPreClose
            dataSet.Tables["2"].Columns[12].ColumnName = "M11"; //SellValue,,,PutSellVol,PutSellVal
            dataSet.Tables["2"].Columns[13].ColumnName = "N11"; //PutBuyVol
            dataSet.Tables["2"].Columns[14].ColumnName = "O11"; //PutBuyVal
            dataSet.Tables["2"].Columns[15].ColumnName = "P11"; //PutSellVol
            dataSet.Tables["2"].Columns[16].ColumnName = "Q11"; //PutSellVal


            return dataSet;
        }
        public DataSet DatTenForeign3(DataSet dataSet)
        {
            // STT,StockCode,BuyVol,BuyValue,SellVol,SellValue
            dataSet.Tables["3"].Columns[1].ColumnName = "B9"; //StockCode
            dataSet.Tables["3"].Columns[2].ColumnName = "C9"; //BuyVol
            dataSet.Tables["3"].Columns[3].ColumnName = "D9"; //BuyValue
            dataSet.Tables["3"].Columns[4].ColumnName = "E9"; //SellVol
            dataSet.Tables["3"].Columns[5].ColumnName = "F9"; //SellValue



            return dataSet;
        }
        public DataSet DatTenForeignCW(DataSet dataSet)
        {

            //STT,StockCode,TotalRoom,CurrentRoom,ForeignOwnedRatio,StateOwnedRadio,BuyVolPreOpen,BuyVolCount,BuyVolPreClose,BuyValue,SellVolPreOpen,SellVolCount,SellVolPreClose,SellValue,PutBuyVol,PutBuyVal,PutSellVol,PutSellVal"
            dataSet.Tables["CW"].Columns[1].ColumnName = "B11"; //StockCode
            dataSet.Tables["CW"].Columns[2].ColumnName = "C11"; //TotalRoom
            dataSet.Tables["CW"].Columns[3].ColumnName = "D11"; //CurrentRoom
            dataSet.Tables["CW"].Columns[4].ColumnName = "E11"; //ForeignOwnedRatio
            dataSet.Tables["CW"].Columns[5].ColumnName = "F11"; //StateOwnedRadio
            dataSet.Tables["CW"].Columns[6].ColumnName = "G11"; //BuyVolPreOpen
            dataSet.Tables["CW"].Columns[7].ColumnName = "H11"; //BuyVolCount
            dataSet.Tables["CW"].Columns[8].ColumnName = "I11"; //BuyVolPreClose
            dataSet.Tables["CW"].Columns[9].ColumnName = "J11"; //BuyValue
            dataSet.Tables["CW"].Columns[10].ColumnName = "K11"; //SellVolPreOpen
            dataSet.Tables["CW"].Columns[11].ColumnName = "L11"; //SellVolCount
            dataSet.Tables["CW"].Columns[12].ColumnName = "M11"; //SellVolPreClose
            dataSet.Tables["CW"].Columns[13].ColumnName = "N11"; //SellValue
            dataSet.Tables["CW"].Columns[14].ColumnName = "O11"; //PutBuyVol
            dataSet.Tables["CW"].Columns[15].ColumnName = "P11"; //PutBuyVal
            dataSet.Tables["CW"].Columns[16].ColumnName = "Q11"; //PutSellVol
            dataSet.Tables["CW"].Columns[17].ColumnName = "R11"; //PutSellVal

            return dataSet;
        }
        public DataSet DatTenForeignODD(DataSet dataSet)
        {
            // STT,StockCode,OrderBuyVol,OrderSellVol,OrderBuyVal,OrderSellVal,PutBuyVol,PutSellVol,PutBuyVal,PutSellVal
            dataSet.Tables["ODD"].Columns[1].ColumnName = "B11"; //StockCode
            dataSet.Tables["ODD"].Columns[2].ColumnName = "C11"; //OrderBuyVol
            dataSet.Tables["ODD"].Columns[3].ColumnName = "D11"; //OrderSellVol
            dataSet.Tables["ODD"].Columns[4].ColumnName = "E11"; //OrderBuyVal
            dataSet.Tables["ODD"].Columns[5].ColumnName = "F11"; //OrderSellVal
            dataSet.Tables["ODD"].Columns[6].ColumnName = "G11"; //PutBuyVol
            dataSet.Tables["ODD"].Columns[7].ColumnName = "H11"; //PutSellVol
            dataSet.Tables["ODD"].Columns[8].ColumnName = "I11"; //PutBuyVal
            dataSet.Tables["ODD"].Columns[9].ColumnName = "J11"; //PutSellVal

            return dataSet;
        }
        public DataSet DatTenForeign5(DataSet dataSet)
        {
            // STT,StockCode,OrderBuyVol,OrderSellVol,OrderBuyVal,OrderSellVal,PutBuyVol,PutSellVol,PutBuyVal,PutSellVal
            dataSet.Tables["5"].Columns[0].ColumnName = "A7"; //StockCode
            dataSet.Tables["5"].Columns[1].ColumnName = "B7"; //OrderBuyVol
            dataSet.Tables["5"].Columns[2].ColumnName = "C7"; //OrderSellVol
            dataSet.Tables["5"].Columns[3].ColumnName = "D7"; //OrderBuyVal
            dataSet.Tables["5"].Columns[4].ColumnName = "E7"; //OrderSellVal
            dataSet.Tables["5"].Columns[5].ColumnName = "F7"; //PutBuyVol
            dataSet.Tables["5"].Columns[6].ColumnName = "G7"; //PutSellVol
            dataSet.Tables["5"].Columns[7].ColumnName = "H7"; //PutBuyVal
            dataSet.Tables["5"].Columns[8].ColumnName = "I7"; //PutSellVal
            dataSet.Tables["5"].Columns[9].ColumnName = "J7"; //PutSellVal
            dataSet.Tables["5"].Columns[10].ColumnName = "K7"; //PutSellVal
            dataSet.Tables["5"].Columns[11].ColumnName = "L7"; //PutSellVal
            dataSet.Tables["5"].Columns[12].ColumnName = "M7"; //PutSellVal
            dataSet.Tables["5"].Columns[13].ColumnName = "N7"; //PutSellVal
            dataSet.Tables["5"].Columns[14].ColumnName = "O7"; //PutSellVal



            return dataSet;
        }
        public DataSet DatTenForeign6(DataSet dataSet)
        {
            // STT,StockCode,BuyVol,BuyValue,SellVol,SellValue
            dataSet.Tables["6"].Columns[1].ColumnName = "B4"; //StockCode
            dataSet.Tables["6"].Columns[2].ColumnName = "C4"; //BuyVol
            dataSet.Tables["6"].Columns[3].ColumnName = "D4"; //BuyValue

            return dataSet;
        }
        //TKGD Tu doanh (Proprietary Trading)
        public DataSet DatTenProprietarySummary(DataSet dataSet)
        {
            // Type,Tradingvolume,rateEntireMaket,TradingValue,ratEntireMaket2
            dataSet.Tables["Tong hop (Summary)"].Columns[1].ColumnName = "B11"; //Tradingvolume
            dataSet.Tables["Tong hop (Summary)"].Columns[2].ColumnName = "C11"; //rateEntireMaket
            dataSet.Tables["Tong hop (Summary)"].Columns[3].ColumnName = "D11"; //TradingValue
            dataSet.Tables["Tong hop (Summary)"].Columns[4].ColumnName = "E11"; //ratEntireMaket2

            return dataSet;
        }
        public DataSet DatTenProprietaryDetails(DataSet dataSet)
        {
            dataSet.Tables["Chi tiet (Details)"].Columns[1].ColumnName = "B13"; //StockCode
            dataSet.Tables["Chi tiet (Details)"].Columns[2].ColumnName = "C13"; //OrderBuyVol
            dataSet.Tables["Chi tiet (Details)"].Columns[3].ColumnName = "D13"; //OrderBuyRateVol
            dataSet.Tables["Chi tiet (Details)"].Columns[4].ColumnName = "E13"; //OrderSellVol
            dataSet.Tables["Chi tiet (Details)"].Columns[5].ColumnName = "F13"; //OrderSellRateVol
            dataSet.Tables["Chi tiet (Details)"].Columns[6].ColumnName = "G13"; //OrderBuyVal
            dataSet.Tables["Chi tiet (Details)"].Columns[7].ColumnName = "H13"; //OrderBuyRateVal
            dataSet.Tables["Chi tiet (Details)"].Columns[8].ColumnName = "I13"; //OrderSellVal
            dataSet.Tables["Chi tiet (Details)"].Columns[9].ColumnName = "J13"; //OrderSellRateVal
            dataSet.Tables["Chi tiet (Details)"].Columns[10].ColumnName = "K13"; //PutBuyVol
            dataSet.Tables["Chi tiet (Details)"].Columns[11].ColumnName = "L13"; //PutBuyRateVol
            dataSet.Tables["Chi tiet (Details)"].Columns[12].ColumnName = "M13"; //PutSellVol
            dataSet.Tables["Chi tiet (Details)"].Columns[13].ColumnName = "N13"; //PutSellRateVol
            dataSet.Tables["Chi tiet (Details)"].Columns[14].ColumnName = "O13"; //PutBuyVal
            dataSet.Tables["Chi tiet (Details)"].Columns[15].ColumnName = "P13"; //PutBuyRateVal
            dataSet.Tables["Chi tiet (Details)"].Columns[16].ColumnName = "Q13"; //PutSellVal
            dataSet.Tables["Chi tiet (Details)"].Columns[17].ColumnName = "R13"; //PutSellRateVal


            return dataSet;
        }
        public DataSet DatTenProprietaryDetailsCW(DataSet dataSet)
        {
            dataSet.Tables["Chi tiet (Details) CW"].Columns[1].ColumnName = "B13"; //StockCode
            dataSet.Tables["Chi tiet (Details) CW"].Columns[2].ColumnName = "C13"; //OrderBuyVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[3].ColumnName = "D13"; //OrderBuyRateVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[4].ColumnName = "E13"; //OrderSellVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[5].ColumnName = "F13"; //OrderSellRateVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[6].ColumnName = "G13"; //OrderBuyVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[7].ColumnName = "H13"; //OrderBuyRateVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[8].ColumnName = "I13"; //OrderSellVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[9].ColumnName = "J13"; //OrderSellRateVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[10].ColumnName = "K13"; //PutBuyVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[11].ColumnName = "L13"; //PutBuyRateVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[12].ColumnName = "M13"; //PutSellVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[13].ColumnName = "N13"; //PutSellRateVol
            dataSet.Tables["Chi tiet (Details) CW"].Columns[14].ColumnName = "O13"; //PutBuyVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[15].ColumnName = "P13"; //PutBuyRateVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[16].ColumnName = "Q13"; //PutSellVal
            dataSet.Tables["Chi tiet (Details) CW"].Columns[17].ColumnName = "R13"; //PutSellRateVal

            return dataSet;
        }
        public DataSet DatTenProprietaryOrder(DataSet dataSet)
        {
            // TransDate,CreateDate,StockCode,BuyOrder,BuyVol,SellOrder,SellVol,BuySellVol
            dataSet.Tables["Dat lenh (Order)"].Columns[1].ColumnName = "B11"; //StockCode
            dataSet.Tables["Dat lenh (Order)"].Columns[2].ColumnName = "C11"; //BuyOrder
            dataSet.Tables["Dat lenh (Order)"].Columns[3].ColumnName = "D11"; //BuyVol
            dataSet.Tables["Dat lenh (Order)"].Columns[4].ColumnName = "E11"; //SellOrder
            dataSet.Tables["Dat lenh (Order)"].Columns[5].ColumnName = "F11"; //SellVol
            dataSet.Tables["Dat lenh (Order)"].Columns[6].ColumnName = "G11"; //BuySellVol

            return dataSet;
        }
        public DataSet DatTenProprietaryOrderCW(DataSet dataSet)
        {
            dataSet.Tables["Dat lenh (Order) CW"].Columns[1].ColumnName = "B11"; //StockCode
            dataSet.Tables["Dat lenh (Order) CW"].Columns[2].ColumnName = "C11"; //BuyOrder
            dataSet.Tables["Dat lenh (Order) CW"].Columns[3].ColumnName = "D11"; //BuyVol
            dataSet.Tables["Dat lenh (Order) CW"].Columns[4].ColumnName = "E11"; //SellOrder
            dataSet.Tables["Dat lenh (Order) CW"].Columns[5].ColumnName = "F11"; //SellVol
            dataSet.Tables["Dat lenh (Order) CW"].Columns[6].ColumnName = "G11"; //BuySellVol
            return dataSet;
        }
        //TK su kien doanh nghiep (Corporate actions)

        public DataSet DatTenCorporate(DataSet dataSet)
        {
            dataSet.Tables["su kien"].Columns[1].ColumnName = "B7"; //StockCode
            dataSet.Tables["su kien"].Columns[2].ColumnName = "C7"; //SectionCode
            dataSet.Tables["su kien"].Columns[3].ColumnName = "D7"; //OutstandingShare
            dataSet.Tables["su kien"].Columns[4].ColumnName = "E7"; //TypeOfAction
            dataSet.Tables["su kien"].Columns[5].ColumnName = "F7"; //ExDate
            dataSet.Tables["su kien"].Columns[6].ColumnName = "G7"; //OfferPrice
            dataSet.Tables["su kien"].Columns[7].ColumnName = "H7"; //ExerciseRatio
            dataSet.Tables["su kien"].Columns[8].ColumnName = "I7"; //RatioForAdjustedPrice
            dataSet.Tables["su kien"].Columns[9].ColumnName = "J7"; //PriorDayClose
            dataSet.Tables["su kien"].Columns[10].ColumnName = "K7"; //RefPriceofExDate
            dataSet.Tables["su kien"].Columns[11].ColumnName = "L7"; //OutstandingShareAfterTheAdjustion

            return dataSet;
        }
        public DataSet DatTenVNAll(DataSet dataSet)
        {
            dataSet.Tables["VNAll"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VNAll"].Columns[2].ColumnName = "C8"; //TodayClose
            dataSet.Tables["VNAll"].Columns[3].ColumnName = "D8"; //OutstandingShares
            dataSet.Tables["VNAll"].Columns[4].ColumnName = "E8"; //[ShareRestrictedOnTransfer]
            dataSet.Tables["VNAll"].Columns[5].ColumnName = "F8"; //[FreeFloat]
            dataSet.Tables["VNAll"].Columns[6].ColumnName = "G8"; //[CapRatio]
            dataSet.Tables["VNAll"].Columns[7].ColumnName = "H8"; //[FreeFloatAdjustedMarketCap]
            dataSet.Tables["VNAll"].Columns[8].ColumnName = "I8"; //[Weight]


            return dataSet;
        }
        public DataSet DatTenVNAll1(DataSet dataSet)
        {
            dataSet.Tables["VNAll"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VNAll"].Columns[2].ColumnName = "C8"; //TodayClose
            dataSet.Tables["VNAll"].Columns[3].ColumnName = "D8"; //OutstandingShares
            dataSet.Tables["VNAll"].Columns[4].ColumnName = "E8"; //[ShareRestrictedOnTransfer]
            dataSet.Tables["VNAll"].Columns[5].ColumnName = "F8"; //[FreeFloat]
            dataSet.Tables["VNAll"].Columns[6].ColumnName = "G8"; //[CapRatio]
            dataSet.Tables["VNAll"].Columns[7].ColumnName = "H8"; //[FreeFloatAdjustedMarketCap]


            return dataSet;
        }
        public DataSet DatTenVN30(DataSet dataSet)
        {
            dataSet.Tables["VN30"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VN30"].Columns[2].ColumnName = "C8"; //Close
            dataSet.Tables["VN30"].Columns[3].ColumnName = "D8"; //OutstandingShares

            return dataSet;
        }
        public DataSet DatTenVNMIDCAP(DataSet dataSet)
        {
            dataSet.Tables["VNMIDCAP"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VNMIDCAP"].Columns[2].ColumnName = "C8"; //Close
            dataSet.Tables["VNMIDCAP"].Columns[3].ColumnName = "D8"; //OutstandingShares

            return dataSet;
        }
        public DataSet DatTenVNSMALLCAP(DataSet dataSet)
        {
            dataSet.Tables["VNSMALLCAP"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VNSMALLCAP"].Columns[2].ColumnName = "C8"; //Close
            dataSet.Tables["VNSMALLCAP"].Columns[3].ColumnName = "D8"; //OutstandingShares

            return dataSet;
        }
        public DataSet DatTenVN100(DataSet dataSet)
        {
            dataSet.Tables["VN100"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VN100"].Columns[2].ColumnName = "C8"; //Close
            dataSet.Tables["VN100"].Columns[3].ColumnName = "D8"; //OutstandingShares

            return dataSet;
        }
        public DataSet DatTenVNALLSHARE(DataSet dataSet)
        {
            dataSet.Tables["VNALLSHARE"].Columns[1].ColumnName = "B8"; //StockCode
            dataSet.Tables["VNALLSHARE"].Columns[2].ColumnName = "C8"; //Close
            dataSet.Tables["VNALLSHARE"].Columns[3].ColumnName = "D8"; //OutstandingShares

            return dataSet;
        }
        public DataSet DatTenTrading_Result0(DataSet dataSet)
        {
            dataSet.Tables["0"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["0"].Columns[3].ColumnName = "D6"; //OrderBuyVol
            dataSet.Tables["0"].Columns[4].ColumnName = "E6"; //OrderBuyRateVol
            dataSet.Tables["0"].Columns[10].ColumnName = "K6"; //OrderSellRateVol
            dataSet.Tables["0"].Columns[11].ColumnName = "L6"; //OrderBuyVal
            dataSet.Tables["0"].Columns[12].ColumnName = "M6"; //OrderBuyRateVal
            dataSet.Tables["0"].Columns[13].ColumnName = "N6"; //OrderSellVal
            dataSet.Tables["0"].Columns[16].ColumnName = "Q6"; //PutBuyRateVol
            dataSet.Tables["0"].Columns[18].ColumnName = "S6"; //PutSellVol
            dataSet.Tables["0"].Columns[19].ColumnName = "T6"; //PutSellRateVol
            dataSet.Tables["0"].Columns[20].ColumnName = "U6"; //PutBuyVal
            dataSet.Tables["0"].Columns[21].ColumnName = "V6"; //PutBuyRateVal
            dataSet.Tables["0"].Columns[22].ColumnName = "W6"; //PutSellVal
            dataSet.Tables["0"].Columns[23].ColumnName = "X6"; //PutSellRateVal
            dataSet.Tables["0"].Columns[24].ColumnName = "Y6"; //PutSellRateVal
            dataSet.Tables["0"].Columns[25].ColumnName = "Z6"; //PutSellRateVal
            dataSet.Tables["0"].Columns[26].ColumnName = "AA6"; //PutSellRateVal

            return dataSet;
        }
        public DataSet DatTenTrading_Summary0(DataSet dataSet)
        {
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[3].ColumnName = "D6"; //OrderBuyVol
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[4].ColumnName = "E6"; //OrderBuyRateVol
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[10].ColumnName = "K6"; //OrderSellRateVol
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[11].ColumnName = "L6"; //OrderBuyVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[12].ColumnName = "M6"; //OrderBuyRateVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[13].ColumnName = "N6"; //OrderSellVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[16].ColumnName = "Q6"; //PutBuyRateVol
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[18].ColumnName = "S6"; //PutSellVol
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[19].ColumnName = "T6"; //PutSellRateVol
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[20].ColumnName = "U6"; //PutBuyVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[21].ColumnName = "V6"; //PutBuyRateVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[22].ColumnName = "W6"; //PutSellVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[23].ColumnName = "X6"; //PutSellRateVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[24].ColumnName = "Y6"; //PutSellRateVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[25].ColumnName = "Z6"; //PutSellRateVal
            dataSet.Tables["OM-CCQ (IFC'S)"].Columns[26].ColumnName = "AA6"; //PutSellRateVal

            return dataSet;
        }

        public DataSet DatTenTrading_Result1(DataSet dataSet)
        {
            dataSet.Tables["1"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["1"].Columns[3].ColumnName = "D6"; //OrderBuyVol
            dataSet.Tables["1"].Columns[4].ColumnName = "E6"; //OrderBuyRateVol
            dataSet.Tables["1"].Columns[10].ColumnName = "K6"; //OrderSellRateVol
            dataSet.Tables["1"].Columns[11].ColumnName = "L6"; //OrderBuyVal
            dataSet.Tables["1"].Columns[12].ColumnName = "M6"; //OrderBuyRateVal
            dataSet.Tables["1"].Columns[13].ColumnName = "N6"; //OrderSellVal
            dataSet.Tables["1"].Columns[16].ColumnName = "Q6"; //PutBuyRateVol
            dataSet.Tables["1"].Columns[18].ColumnName = "S6"; //PutSellVol
            dataSet.Tables["1"].Columns[19].ColumnName = "T6"; //PutSellRateVol
            dataSet.Tables["1"].Columns[20].ColumnName = "U6"; //PutBuyVal
            dataSet.Tables["1"].Columns[21].ColumnName = "V6"; //PutBuyRateVal
            dataSet.Tables["1"].Columns[22].ColumnName = "W6"; //PutSellVal
            dataSet.Tables["1"].Columns[23].ColumnName = "X6"; //PutSellRateVal
            dataSet.Tables["1"].Columns[24].ColumnName = "Y6"; //PutSellRateVal
            dataSet.Tables["1"].Columns[25].ColumnName = "Z6"; //PutSellRateVal
            dataSet.Tables["1"].Columns[26].ColumnName = "AA6"; //PutSellRateVal

            return dataSet;
        }
        public DataSet DatTenTrading_Summary1(DataSet dataSet)
        {
            dataSet.Tables["OM-CP(Stocks)"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["OM-CP(Stocks)"].Columns[3].ColumnName = "D6"; //OrderBuyVol
            dataSet.Tables["OM-CP(Stocks)"].Columns[4].ColumnName = "E6"; //OrderBuyRateVol
            dataSet.Tables["OM-CP(Stocks)"].Columns[10].ColumnName = "K6"; //OrderSellRateVol
            dataSet.Tables["OM-CP(Stocks)"].Columns[11].ColumnName = "L6"; //OrderBuyVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[12].ColumnName = "M6"; //OrderBuyRateVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[13].ColumnName = "N6"; //OrderSellVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[16].ColumnName = "Q6"; //PutBuyRateVol
            dataSet.Tables["OM-CP(Stocks)"].Columns[18].ColumnName = "S6"; //PutSellVol
            dataSet.Tables["OM-CP(Stocks)"].Columns[19].ColumnName = "T6"; //PutSellRateVol
            dataSet.Tables["OM-CP(Stocks)"].Columns[20].ColumnName = "U6"; //PutBuyVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[21].ColumnName = "V6"; //PutBuyRateVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[22].ColumnName = "W6"; //PutSellVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[23].ColumnName = "X6"; //PutSellRateVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[24].ColumnName = "Y6"; //PutSellRateVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[25].ColumnName = "Z6"; //PutSellRateVal
            dataSet.Tables["OM-CP(Stocks)"].Columns[26].ColumnName = "AA6"; //PutSellRateVal

            return dataSet;
        }
        public DataSet DatTenTrading_Result2(DataSet dataSet)
        {
            dataSet.Tables["2"].Columns[0].ColumnName = "A5"; //StockCode
            dataSet.Tables["2"].Columns[1].ColumnName = "B5"; //OrderBuyVol
            dataSet.Tables["2"].Columns[2].ColumnName = "C5"; //OrderBuyRateVol
            dataSet.Tables["2"].Columns[3].ColumnName = "D5"; //OrderSellRateVol
            dataSet.Tables["2"].Columns[4].ColumnName = "E5"; //OrderBuyVal
            dataSet.Tables["2"].Columns[5].ColumnName = "F5"; //OrderBuyRateVal
            dataSet.Tables["2"].Columns[6].ColumnName = "G5"; //OrderSellVal
            dataSet.Tables["2"].Columns[7].ColumnName = "H5"; //PutBuyRateVol            dataSet.Tables["2"].Columns[7].ColumnName = "H5"; //PutBuyRateVol
            dataSet.Tables["2"].Columns[8].ColumnName = "I5"; //PutBuyRateVol
            dataSet.Tables["2"].Columns[9].ColumnName = "J5"; //PutBuyRateVol
            dataSet.Tables["2"].Columns[10].ColumnName = "K5"; //PutBuyRateVol
            return dataSet;
        }
        public DataSet DatTenTrading_Summary2(DataSet dataSet)
        {
            dataSet.Tables["PT-TP(Bonds)"].Columns[0].ColumnName = "A5"; //StockCode
            dataSet.Tables["PT-TP(Bonds)"].Columns[1].ColumnName = "B5"; //OrderBuyVol
            dataSet.Tables["PT-TP(Bonds)"].Columns[2].ColumnName = "C5"; //OrderBuyRateVol
            dataSet.Tables["PT-TP(Bonds)"].Columns[3].ColumnName = "D5"; //OrderSellRateVol
            dataSet.Tables["PT-TP(Bonds)"].Columns[4].ColumnName = "E5"; //OrderBuyVal
            dataSet.Tables["PT-TP(Bonds)"].Columns[5].ColumnName = "F5"; //OrderBuyRateVal
            dataSet.Tables["PT-TP(Bonds)"].Columns[6].ColumnName = "G5"; //OrderSellVal
            dataSet.Tables["PT-TP(Bonds)"].Columns[7].ColumnName = "H5"; //PutBuyRateVol            dataSet.Tables["PT-TP(Bonds)"].Columns[7].ColumnName = "H5"; //PutBuyRateVol
            dataSet.Tables["PT-TP(Bonds)"].Columns[8].ColumnName = "I5"; //PutBuyRateVol
            dataSet.Tables["PT-TP(Bonds)"].Columns[9].ColumnName = "J5"; //PutBuyRateVol
            dataSet.Tables["PT-TP(Bonds)"].Columns[10].ColumnName = "K5"; //PutBuyRateVol
            return dataSet;
        }
        public DataSet DatTenTrading_Result3(DataSet dataSet)
        {
            dataSet.Tables["3"].Columns[0].ColumnName = "A5"; //StockCode
            dataSet.Tables["3"].Columns[1].ColumnName = "B5"; //OrderBuyVol
            dataSet.Tables["3"].Columns[2].ColumnName = "C5"; //OrderBuyRateVol
            dataSet.Tables["3"].Columns[4].ColumnName = "E5"; //OrderSellRateVol
            dataSet.Tables["3"].Columns[5].ColumnName = "F5"; //OrderBuyVal
            dataSet.Tables["3"].Columns[6].ColumnName = "G5"; //OrderBuyRateVal

            return dataSet;
        }
        public DataSet DatTenTrading_Summary3(DataSet dataSet)
        {
            dataSet.Tables["PT-CP&CCQ(Stocks & IFCs)"].Columns[0].ColumnName = "A5"; //StockCode
            dataSet.Tables["PT-CP&CCQ(Stocks & IFCs)"].Columns[1].ColumnName = "B5"; //OrderBuyVol
            dataSet.Tables["PT-CP&CCQ(Stocks & IFCs)"].Columns[2].ColumnName = "C5"; //OrderBuyRateVol
            dataSet.Tables["PT-CP&CCQ(Stocks & IFCs)"].Columns[4].ColumnName = "E5"; //OrderSellRateVol
            dataSet.Tables["PT-CP&CCQ(Stocks & IFCs)"].Columns[5].ColumnName = "F5"; //OrderBuyVal
            dataSet.Tables["PT-CP&CCQ(Stocks & IFCs)"].Columns[6].ColumnName = "G5"; //OrderBuyRateVal

            return dataSet;
        }
        public DataSet DatTenTrading_Result4(DataSet dataSet)
        {
            dataSet.Tables["4"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["4"].Columns[2].ColumnName = "C6"; //OrderBuyVol
            dataSet.Tables["4"].Columns[3].ColumnName = "D6"; //OrderBuyRateVol
            dataSet.Tables["4"].Columns[4].ColumnName = "E6"; //OrderSellRateVol
            dataSet.Tables["4"].Columns[5].ColumnName = "F6"; //OrderBuyVal
            dataSet.Tables["4"].Columns[6].ColumnName = "G6"; //OrderBuyRateVal
            dataSet.Tables["4"].Columns[7].ColumnName = "H6"; //OrderBuyRateVal
            dataSet.Tables["4"].Columns[8].ColumnName = "I6"; //OrderBuyRateVal
            dataSet.Tables["4"].Columns[9].ColumnName = "J6"; //OrderBuyRateVal
            dataSet.Tables["4"].Columns[10].ColumnName = "K6"; //OrderBuyRateVal

            return dataSet;
        }
        public DataSet DatTenTrading_Summary4(DataSet dataSet)
        {
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[1].ColumnName = "B6"; //StockCode
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[2].ColumnName = "C6"; //OrderBuyVol
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[3].ColumnName = "D6"; //OrderBuyRateVol
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[4].ColumnName = "E6"; //OrderSellRateVol
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[5].ColumnName = "F6"; //OrderBuyVal
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[6].ColumnName = "G6"; //OrderBuyRateVal
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[7].ColumnName = "H6"; //OrderBuyRateVal
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[8].ColumnName = "I6"; //OrderBuyRateVal
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[9].ColumnName = "J6"; //OrderBuyRateVal
            dataSet.Tables["CPQ (Treasury stocks)"].Columns[10].ColumnName = "K6"; //OrderBuyRateVal

            return dataSet;
        }
        public DataSet DatTenPutThoughPT(DataSet dataSet)
        {
            dataSet.Tables["PT"].Columns[1].ColumnName = "B7"; //StockCode
            dataSet.Tables["PT"].Columns[2].ColumnName = "C7"; //OrderBuyVol
            dataSet.Tables["PT"].Columns[3].ColumnName = "D7"; //OrderBuyRateVol
            dataSet.Tables["PT"].Columns[6].ColumnName = "G7"; //OrderSellRateVol
            dataSet.Tables["PT"].Columns[7].ColumnName = "H7"; //OrderBuyVal

            return dataSet;
        }
        public DataSet DatTenPutThoughMBL(DataSet dataSet)
        {
            dataSet.Tables["MBL"].Columns[1].ColumnName = "B9"; //StockCode
            dataSet.Tables["MBL"].Columns[2].ColumnName = "C9"; //OrderBuyVol
            dataSet.Tables["MBL"].Columns[3].ColumnName = "D9"; //OrderBuyRateVol
            dataSet.Tables["MBL"].Columns[4].ColumnName = "E9"; //OrderSellRateVol
            dataSet.Tables["MBL"].Columns[5].ColumnName = "F9"; //OrderBuyVal
            dataSet.Tables["MBL"].Columns[6].ColumnName = "G9"; //OrderBuyRateVal
            dataSet.Tables["MBL"].Columns[7].ColumnName = "H9"; //OrderBuyRateVal
            dataSet.Tables["MBL"].Columns[8].ColumnName = "I9"; //OrderBuyRateVal
            dataSet.Tables["MBL"].Columns[9].ColumnName = "J9"; //OrderBuyRateVal
            dataSet.Tables["MBL"].Columns[10].ColumnName = "K9"; //OrderBuyRateVal

            return dataSet;
        }
        public DataSet DatTenMatketCap(DataSet dataSet)
        {
            dataSet.Tables["1"].Columns[0].ColumnName = "A7"; //StockCode
            dataSet.Tables["1"].Columns[1].ColumnName = "B7"; //StockCode
            dataSet.Tables["1"].Columns[2].ColumnName = "C7"; //OrderBuyVol
            dataSet.Tables["1"].Columns[3].ColumnName = "D7"; //OrderBuyRateVol
            dataSet.Tables["1"].Columns[4].ColumnName = "E7"; //OrderSellRateVol
            dataSet.Tables["1"].Columns[5].ColumnName = "F7"; //OrderBuyVal
            dataSet.Tables["1"].Columns[6].ColumnName = "G7"; //OrderBuyRateVal
            dataSet.Tables["1"].Columns[7].ColumnName = "H7"; //OrderBuyRateVal
            dataSet.Tables["1"].Columns[8].ColumnName = "I7"; //OrderBuyRateVal
            dataSet.Tables["1"].Columns[9].ColumnName = "J7"; //OrderBuyRateVal
            dataSet.Tables["1"].Columns[10].ColumnName = "K7"; //OrderBuyRateVal

            return dataSet;
        }
        public DataSet DatTenOrrderMatching(DataSet dataSet)
        {
            dataSet.Tables["2"].Columns[0].ColumnName = "A9"; //StockCode
            dataSet.Tables["2"].Columns[1].ColumnName = "B9"; //StockCode
            dataSet.Tables["2"].Columns[2].ColumnName = "C9"; //OrderBuyVol
            dataSet.Tables["2"].Columns[3].ColumnName = "D9"; //OrderBuyRateVol
            dataSet.Tables["2"].Columns[4].ColumnName = "E9"; //OrderSellRateVol
            dataSet.Tables["2"].Columns[5].ColumnName = "F9"; //OrderBuyVal
            dataSet.Tables["2"].Columns[7].ColumnName = "H9"; //OrderBuyRateVal
            dataSet.Tables["2"].Columns[8].ColumnName = "I9"; //OrderBuyRateVal
            dataSet.Tables["2"].Columns[9].ColumnName = "J9"; //OrderBuyRateVal
            dataSet.Tables["2"].Columns[10].ColumnName = "K9"; //OrderBuyRateVal
            dataSet.Tables["2"].Columns[11].ColumnName = "L9"; //OrderBuyRateVal

            return dataSet;
        }
    }
}