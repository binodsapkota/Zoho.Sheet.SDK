using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Domain
{
    public class Sheet
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    public class WorksheetRecordCriteria
    {
        public string Key { get; set; }
        public string Operator { get; set; }
        public object Matcher { get; set; }
        public string Type { get; set; }
    }
}
