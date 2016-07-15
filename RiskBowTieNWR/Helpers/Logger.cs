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

        public int ErrorCount { get; private set; }

        public Logger(MainViewModel vm)
        {
            _vm = vm;
        }

        public void Log(string message)
        {
            _vm.ShowWaitFormNow(message);
            _vm.ProgressLogText += $"{message}\n";
        }

        public void LogError(string message)
        {
            _vm.ProgressLogText += $"ERROR: {message}\n";
            ErrorCount++;
        }
    }
}
