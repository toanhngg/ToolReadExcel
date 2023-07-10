using System;

namespace ApiExcelToDB.Entities
{
    public class ETableTrading_Result
    {
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public string StockCode { get; set; }

        public float Change { get; set; }
        public float OpenPrice { get; set; }
        public float HighPrice { get; set; }
        public float LowPrice { get; set; }
        public float AvgPrice { get; set; }
        public float ClosePrice { get; set; }
        public float RefPrice { get; set; }
        public float TotalShareVol { get; set; }
        public float TotalValue { get; set; }
        public float BidOrders { get; set; }
        public float OfferOrders { get; set; }
        public float BidVol { get; set; }
        public float OfferVol { get; set; }
        public float RefPriceAfter { get; set; }
        public float CeilingPrice { get; set; }
        public float  FloorPrice { get; set; }


        public float CYield { get; set; }
        public float TradingVolume { get; set; }
        public float TradingValue { get; set; }
        public float MaturityDate { get; set; }
        public float InterestRate { get; set; }
        public string InterestPaymentMethod { get; set; }
        public float TimestoMaturity { get; set; }
        public float YTM { get; set; }

        public string BuySale { get; set; }
        public float RegistrationVol { get; set; }
        public float RateRegistrationVol1 { get; set; }
        public float AccumulatedTradingVol { get; set; }
        public float RateRegistrationVol2 { get; set; }
        public float RemainingVol { get; set; }
        public float RateRegistrationVol3 { get; set; }
        public string Deadline { get; set; }



    }
}
