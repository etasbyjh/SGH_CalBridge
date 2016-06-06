using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Caliburn.Micro;

using System.Windows.Controls;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Threading;
using System.Collections.ObjectModel;

namespace SGH_CalBridge
{
    class SGH_CalBridgeViewModel : PropertyChangedBase
    {

        public SGH_CalBridgeViewModel()
        {
            inputParams = new BindableCollection<SGH_BridgeParameter>();

            inputPath = "Select Your Input Excel File";
            inputSheetNumber = "0";
            inputParamUserName = "Specify Your Input Parameter's User Name";
            inputParamExcelNameBox = "Specify Your Input Parameter's Excel Name Box";
            
            
        }


        private string inputPath;
        public string InputPath
        {
            get { return inputPath; }
            set
            {
                inputPath = value;
                NotifyOfPropertyChange(() => InputPath);
            }
        }

        private string inputSheetNumber;
        public string InputSheetNumber
        {
            get { return inputSheetNumber; }
            set
            {
                inputSheetNumber = value;
                NotifyOfPropertyChange(() => InputSheetNumber);
            }
        }

        private string inputParamUserName;
        public string InputParamUserName
        {
            get { return inputParamUserName; }
            set
            {
                inputParamUserName = value;
                NotifyOfPropertyChange(() => InputParamUserName);
            }
        }

        private string inputParamExcelNameBox;
        public string InputParamExcelNameBox
        {
            get { return inputParamExcelNameBox; }
            set
            {
                inputParamExcelNameBox = value;
                NotifyOfPropertyChange(() => InputParamExcelNameBox);
            }
        }

        private BindableCollection<SGH_BridgeParameter> inputParams;
        public BindableCollection<SGH_BridgeParameter> InputParams
        {
            get { return inputParams; }
            set
            {
                inputParams = value;
                NotifyOfPropertyChange(() => InputParams);
            }
        }

        private SGH_BridgeParameter activeInputParam;
        public SGH_BridgeParameter ActiveInputParam
        {
            get { return activeInputParam; }
            set
            {
                activeInputParam = value;
                NotifyOfPropertyChange(() => ActiveInputParam);
            }
        }

        public void GetInputFilePath()
        {
            OpenFileDialog dig = new OpenFileDialog();
            
            dig.DefaultExt = ".xlsx"; // Default file extension 
            dig.Filter = "Excel workbook (.xlsx)|*.xlsx"; // Filter files by extension 

            Nullable<bool> result = dig.ShowDialog();
            if(result == true)
            {
                InputPath = dig.FileName;
            }
            else
            {
                InputPath = "Select Your Input Excel File";
            }

        }

        public void AddItemToList()
        {
            SGH_BridgeParameter pex = new SGH_BridgeParameter(inputParamExcelNameBox);
            pex.IsInput = true;
            pex.userName = inputParamUserName;
            pex.updatePreivewName();
            
            inputParams.Add(pex);
        }

        public void RemoveSelectedItem()
        {
            if(activeInputParam != null)
            {
                inputParams.Remove(activeInputParam);
            }
        }

    }
}
