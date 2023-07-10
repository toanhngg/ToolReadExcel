using System;

namespace ApiExcelToDB.Entities
{
    public class ETableCC
    {
        public DateTime CreateDate { get; set; }
        public DateTime TransDate { get; set; }
        public string StockCode { get; set; }
        public float BuyingOrders { get; set; }
        public float BuyingVolume { get; set; }
        public float SellingOrders { get; set; }
        public float SellingVolume { get; set; }
        public float Change { get; set; }
        public float CeilingPrice { get; set; }
        public float FloorPrice { get; set; }
        public float BestBidPrice { get; set; }
        public float BestOfferPrice { get; set; }
        public float Spreads { get; set; }
    }
}
