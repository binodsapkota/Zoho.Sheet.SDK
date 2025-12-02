namespace Zoho.Sheet.SDK.Domain
{
    public class Workbook
    {
        public string Id { get; set; }              // maps resource_id
        public string Name { get; set; }            // maps workbook_name
        public string Url { get; set; }             // maps workbook_url
        public string CreatedBy { get; set; }       // maps created_by
        public string CreatedTime { get; set; }     // maps created_time
        public string LastModifiedTime { get; set; } // maps last_modified_time
    }
}
