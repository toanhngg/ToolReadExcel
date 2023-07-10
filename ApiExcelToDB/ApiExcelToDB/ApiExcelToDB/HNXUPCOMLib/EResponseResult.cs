namespace ApiExcelToDB.HNXUPCOMLib
{
    public class EResponseResult
    {
        /// <summary>
        /// code return tu sp hoac system neu co exception
        /// </summary>
        public long Code { get; set; }

        /// <summary>
        /// message return tu sp hoac error msg neu co exception
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// data khong xac dinh type
        /// </summary>
        public object Data { get; set; }

    }
}
