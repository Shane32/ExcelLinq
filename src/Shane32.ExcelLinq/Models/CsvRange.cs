using System;
using System.Collections.Generic;
using System.Text;

namespace Shane32.ExcelLinq.Models
{
    public class CsvRange
    {
        public int StartColumn { get; set; }
        public int StartRow { get; set; }
        public int? EndColumn { get; set; }
        public int? EndRow { get; set; }
    }
}
