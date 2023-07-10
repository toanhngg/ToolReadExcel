namespace ApiExcelToDB.Entities
{
    public class ConfigApp
    {
        public DataView Data { get; set; }

        public struct DataView
        {
            public string SheetName { get; set; }
            public string TableName { get; set; }
            public string SPName { get; set; }

            public string BeginCell { get; set; }
            public string Column { get; set; }

        }
    }
}
