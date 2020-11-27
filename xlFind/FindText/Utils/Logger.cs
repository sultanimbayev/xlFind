using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlFind.Utils
{
    public class Logger
    {
        /// <summary>
        /// Настроики логирования
        /// </summary>
        public LoggerSettings Settings { get; private set; }
        
        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="logFilename"></param>
        public Logger(string logFilename = null)
        {
            Settings = new LoggerSettings();
            Settings.LogFilename = logFilename;
            if(logFilename == null) { Settings.WriteToFile = false; }
        }
        
        /// <summary>
        /// Основная функция для записи лога
        /// </summary>
        /// <param name="msg">сообщение логирования</param>
        public void Log(string msg)
        {
            if (Settings.WriteToFile)
            {
                __createDirectoryIfNotExists(Settings.LogFilename);
                var _logFs = new FileStream(Settings.LogFilename, FileMode.Append, FileAccess.Write);
                var logData = Encoding.UTF8.GetBytes(msg + "\r\n");
                _logFs.Write(logData, 0, logData.Length);
            }
            if (Settings.EchoConsole) { Console.WriteLine(msg); }
        }

        private void __createDirectoryIfNotExists(string filepath)
        {
            if (!Directory.Exists(Path.GetDirectoryName(filepath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filepath));
            }
        }
        

        public class LoggerSettings
        {
            private Action<string> _logFilenameChangedDeleg;

            public LoggerSettings(Action<string> LogFilenameChangedDeleg = null)
            {
                _logFilenameChangedDeleg = LogFilenameChangedDeleg;
            }

            public bool EchoConsole { get; set; } = false;

            public bool WriteToFile { get; set; } = true;

            private string _logFilename;
            public string LogFilename {
                get { return _logFilename; }
                set
                {
                    _logFilenameChangedDeleg?.Invoke(value);
                    _logFilename = value;
                }
            }



        }
    }

    
}
