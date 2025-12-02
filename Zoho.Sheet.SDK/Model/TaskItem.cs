using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Model
{
    public class TaskItem
    {
        public string TaskName { get; set; }
        public string AssignedTo { get; set; }
        public string Status { get; set; }
        public int Priority { get; set; }
        public string DueDate { get; set; } // or DateTime if needed
    }
}
