using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Udilovich.ExcelClient
{
    public class ExcelOutputEntry
    {
        public string Worksheet { get; set; }
        public string[,] Values { get; set; }
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
    }
}
