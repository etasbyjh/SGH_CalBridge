using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Caliburn.Micro;

namespace SGH_CalBridge
{
    class SGH_CalBridgeBootStrapper : BootstrapperBase
    {
        public SGH_CalBridgeBootStrapper()
        {
            Initialize();
        }

        protected override void OnStartup(object sender, StartupEventArgs e)
        {
            DisplayRootViewFor<SGH_CalBridgeViewModel>();
        }
    }
}
