using System;
namespace ApiExcelToDB.Model
{
    public class TinhHinhDatLenh2011
    {
      
        public string Symbol { get; set; }
      
        public double NumberofBids_QT { get; set; }
        public double BidVolume_QT { get; set; }
        public double NumberofOffers_QT { get; set; }
        public double OfferVolume_QT { get; set; }
        public double Difference_QT { get; set; }
        public double NumberofBids_NT { get; set; }
        public double BidVolume_NT { get; set; }
        public double NumberofOffers_NT { get; set; }

        public double OfferVolume_NT { get; set; }

        public double Difference_NT { get; set; }
        public double SLDatMua { get; set; }
        public double KLDatMua { get; set; }
        public double SLDatBan { get; set; }
        public double KLDatBan { get; set; }
        //Symbol,NumberofBids_QT,BidVolume_QT,NumberofOffers_QT,OfferVolume_QT,Difference_QT,NumberofBids_NT
        //,BidVolume_NT,NumberofOffers_NT,OfferVolume_NT,Difference_NT,SLDatMua,KLDatMua,SLDatBan,KLDatBan,Trangding_Date

        public DateTime Trangding_Date { get; set; }
    }
}
