using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Udilovich.ExcelClient
{
    public class ExcelEntry
    {
        public ExcelEntry()
        {

        }
        public ExcelEntry(string ValueName, string ValueMagnitude)
        {
            this.ValueMagnitude = ValueMagnitude;
            this.ValueName = ValueName;
        }
        public string ValueName { get; set; }
        public string ValueMagnitude { get; set; }
    }
}
