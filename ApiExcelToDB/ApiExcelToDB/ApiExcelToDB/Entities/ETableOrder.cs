using System;

namespace ApiExcelToDB.Entities
{
    public class ETableOrder
    {
        public string StockCode { get; set; }
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public float BuyingOrders { get; set; }
        public float BuyingVolume { get; set; }
        public float SellingOrders { get; set; }
        public float SellingVolume { get; set; }
        public float TradingVolume { get; set; }
        public float BuySellVolume { get; set; }


    }
   
}
