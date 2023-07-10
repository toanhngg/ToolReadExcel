namespace ApiExcelToDB.HNXUPCOMLib
{
    public class EDalResult
    {
        public const long __CODE_ERROR = -2; // co loi
        public const long __CODE_INIT = -1; // khoi tao
        public const long __CODE_SUCCESS = 0;    // thanh cong
        public const string __STRING_INIT = "INIT"; // moi khoi tao object xong
        public const string __STRING_ERROR = "ERROR"; // co loi xay ra trong app code
        public const string __STRING_SUCCESS = "SUCCESS"; // thanh cong
        public const string __STRING_GET_SCRIPT = "GetScript";
        public const string __STRING_RETURN_NEW_LINE = "\r\n";
        public const string __STRING_EXEC = "EXEC";
        public const string __STRING_SPACE = " ";
        public const string __STRING_INSERT = "INSERT INTO ";
        public const string __STRING_VALUES = " VALUES ";

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

        /// <summary>
        /// constructor
        /// </summary>
        public EDalResult() { Code = __CODE_INIT; Message = __STRING_INIT; Data = null; }
    }
}
