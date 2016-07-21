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

        private int _textMode = 0;

        public int ErrorCount { get; private set; }

        public Logger(MainViewModel vm, int textMode = 0)
        {
            _vm = vm;
            _textMode = textMode;
        }

        public void Log(string message)
        {
            _vm.ShowWaitFormNow(message);
            SetMessage(message);
        }

        public void LogError(string message)
        {
            SetMessage($"ERROR: {message}\n");
            ErrorCount++;
        }

        public void HideProgress()
        {
            _vm.ShowWaitForm = false;
        }

        private void SetMessage(string message)
        {
            switch (_textMode)
            {
                case 0:
                case 1:
                    _vm.ProgressLogText += $"{message}\n";
                    break;
                case 2:
                    _vm.ProgressLogText2 += $"{message}\n";
                    break;
                default:
                    throw new Exception("Unknown Log type");
            }
        }
    }
}
