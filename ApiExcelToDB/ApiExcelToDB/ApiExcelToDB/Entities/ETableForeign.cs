using System;

namespace ApiExcelToDB.Entities
{
    public class ETableForeign
    {
        //  STT, StockCode, OrderBuyVol, OrderSellVol, OrderBuyVal, OrderSellVal, PutBuyVol, PutSellVol, PutBuyVal, PutSellVal
        public DateTime CreateDate { get; set; }
        public DateTime TransDate { get; set; }
        public string StockCode { get; set; }
        public float OrderBuyVol { get; set; }
        public float OrderSellVol { get; set; }
        public float OrderBuyVal { get; set; }
        public float OrderSellVal { get; set; }
        public float PutBuyVol { get; set; }
        public float PutSellVol { get; set; }
        public float PutBuyVal { get; set; }
        public float PutSellVal { get; set; }

        //STT,StockCode,TotalRoom,CurrentRoom,ForeignOwnedRatio,StateOwnedRadio,BuyVolPreOpen,BuyVolCount,BuyVolPreClose,BuyValue,SellVolPreOpen,SellVolCount,SellVolPreClose,SellValue,PutBuyVol,PutBuyVal,PutSellVol,PutSellVal"
        public float TotalRoom { get; set; }
        public float CurrentRoom { get; set; }
        public float ForeignOwnedRatio { get; set; }
      //  public float StateOwnedRadio { get; set; }
        public float BuyVolPreOpen { get; set; }
        public float BuyVolCount { get; set; }
        public float BuyVolPreClose { get; set; }
        public float BuyValue { get; set; }
        public float SellVolPreOpen { get; set; }
        public float SellVolCount { get; set; }
        public float SellVolPreClose { get; set; }
        public float SellValue { get; set; }
        //    public float PutBuyVol { get; set; }
        //    public float PutBuyVal { get; set; }
        //    public float PutSellVol { get; set; }
        //     public float PutSellVal { get; set; }


        //STT,StockCode,BuyVol,BuyValue,SellVol,SellValue
        public float BuyVol { get; set; }
        //    public float BuyValue { get; set; }
        public float SellVol { get; set; }
        //    public float SellValue { get; set; }
        //Type,Tradingvolume,rateEntireMaket,TradingValue,ratEntireMaket2

        //STT,StockCode,TotalRoom,CurrentRoom,ForeignOwnedRatio,StateOwnedRadio,BuyVolPreOpen,BuyVolCount,BuyVolPreClose,BuyValue,SellVolPreOpen,SellVolCount,SellVolPreClose,SellValue,PutBuyVol,PutBuyVal,PutSellVol,PutSellVal"
        public float Tradingvolume { get; set; }
        public float rateEntireMaket { get; set; }
        public float TradingValue { get; set; }
        public float ratEntireMaket2 { get; set; }
        public string Type { get; set; }



       public float StateOwnedRatio { get; set; }
       public float OrderBuyPreOpen { get; set; }
       public float OrderBuyCont { get; set; }
       public float OrderBuyPreClose { get; set; }
       public float OrderSellPreOpen { get; set; }
       public float OpenSellCont { get; set; }
       public float OpenSellClose { get; set; }
      
       public float TotalBuyVol { get; set; }
       public float TotalSellVol { get; set; }





    }
}
