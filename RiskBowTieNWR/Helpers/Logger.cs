using System;
using System.Collections.Generic;
using System.IO;
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

        private static readonly string _SCNWR = "SharpCloudNWR";

        private readonly string _localPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), _SCNWR);

        public Logger(MainViewModel vm, int textMode = 0)
        {
            Directory.CreateDirectory(_localPath);

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
            SetMessage($"ERROR: {message}");
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
                    break;
                case 1:
                    _vm.ProgressLogText += $"{message}\n";
                    break;
                case 2:
                    _vm.ProgressLogText2 += $"{message}\n";
                    break;
                    break;
                default:
                    throw new Exception("Unknown Log type");
            }
            SaveToLogFile(message);
        }

        private void SaveToLogFile(string message)
        {
            var now = DateTime.UtcNow;
            var path = $"{_localPath}/{now.Year}-{now.Month}-{now.Day}.log";

            File.AppendAllText(path, $"{now.ToLongTimeString()}\t{message}\r\n");
        }

        public static void ShowLogFolder()
        {
            System.Diagnostics.Process.Start("explorer.exe", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), _SCNWR));
        }
    }
}
