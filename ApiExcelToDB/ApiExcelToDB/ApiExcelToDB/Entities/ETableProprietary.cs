using System;

namespace ApiExcelToDB.Entities
{
    public class ETableProprietary
    {
        public DateTime CreateDate { get; set; }
        public DateTime TransDate { get; set; }
        public string StockCode { get; set; }
        public float Tradingvolume { get; set; }
        public float rateEntireMaket { get; set; }
        public float TradingValue { get; set; }
        public float ratEntireMaket2 { get; set; }
        public string Type { get; set; }
        //
        public float OrderBuyVol { get; set; }
        public float OrderBuyRateVol { get; set; }
        public float OrderSellVol { get; set; }
        public float OrderSellRateVol { get; set; }
        public float OrderBuyVal { get; set; }
        public float OrderBuyRateVal { get; set; }
        public float OrderSellVal { get; set; }
        public float OrderSellRateVal { get; set; }
        public float PutBuyVol { get; set; }
        public float PutBuyRateVol { get; set; }
        public float PutSellVol { get; set; }
        public float PutSellRateVol { get; set; }
        public float PutBuyVal { get; set; }
        public float PutBuyRateVal { get; set; }
        public float PutSellVal { get; set; }
        public float PutSellRateVal { get; set; }

        public float BuyOrder { get; set; }
        public float BuyVol { get; set; }
        public float SellOrder { get; set; }
        public float SellVol { get; set; }
        public float BuySellVol { get; set; }

    }
}
