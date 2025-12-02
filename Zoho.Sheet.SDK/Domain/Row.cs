using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Domain
{
    public class Row
    {
        public int RowIndex { get; set; }
        public List<object> Values { get; set; }
    }
}
