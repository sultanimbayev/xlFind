using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Models
{
    public class ValueReplacedLog
    {
        public string Filename { get; set; }
        public string WorksheetName { get; set; }
        public string CellAddress { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
}
