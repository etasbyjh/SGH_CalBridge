using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Udilovich.ExcelClient
{
    public class ExcelAddress
    {
        public int Row { get; set; }
        public int Column { get; set; }

        public ExcelAddress(int Row, int Column)
        {
            this.Row = Row;
            this.Column=Column ;
        }
    }
}
