using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGH_CalBridge
{
    public class SGH_BridgeParameter
    {
        public int ID { get; set; }
        public string previewName { get; set; }
        public string userName { get; set; }
        public string excelNameBox { get; set; }
        public bool IsInput { get; set; }
        
        public SGH_BridgeParameter(string boxName)
        {
            this.excelNameBox = boxName;
            this.previewName = "";
        }

        public void updatePreivewName()
        {
            this.previewName = "User Input Paramter: " + this.userName + ", Excel NameBox: " + this.excelNameBox;
        }
    }
}
