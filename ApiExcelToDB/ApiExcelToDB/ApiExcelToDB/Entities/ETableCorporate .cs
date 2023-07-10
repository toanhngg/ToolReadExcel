using System;

namespace ApiExcelToDB.Entities
{
    public class ETableCorporate
    {
        public DateTime TransDate { get; set; }
        public DateTime CreateDate { get; set; }
        public string StockCode { get; set; }
        public float SectionCode { get; set; }
        public float OutstandingShare { get; set; }
        public string TypeOfAction { get; set; }
        public string ExDate { get; set; }
        public float OfferPrice { get; set; }
        public string ExerciseRatio { get; set; }
        public float RatioForAdjustedPrice { get; set; }
        public float PriorDayClose { get; set; }
        public float RefPriceofExDate { get; set; }
        public float OutstandingShareAfterTheAdjustion { get; set; }
    }
}
