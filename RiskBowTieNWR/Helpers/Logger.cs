using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using RiskBowTieNWR.ViewModels;

namespace RiskBowTieNWR.Helpers
{
    public class Logger
    {
        private MainViewModel _vm;

        public Logger(MainViewModel vm)
        {
            _vm = vm;
        }

        public void Log(string message)
        {
            _vm.ShowWaitFormNow(message);

            _vm.ProgressLogText += $"{message}\n";
        }
    }
}
