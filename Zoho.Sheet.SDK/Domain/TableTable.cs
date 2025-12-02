namespace Zoho.Sheet.SDK.Domain
{
    public class Table
    {
        public string TableName { get; set; }
        public int TableId { get; set; }
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int EndRow { get; set; }
        public int EndColumn { get; set; }
    }
    public class TableRecord
    {
        public Dictionary<string, object> Data { get; set; } = new();
        public int RowIndex { get; set; }
    }

    public class TableRecordCriteria
    {
        public string Key { get; set; }
        public string Operator { get; set; }
        public object Matcher { get; set; }
        public string Type { get; set; }
    }

}
