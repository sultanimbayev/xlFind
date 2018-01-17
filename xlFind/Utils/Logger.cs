using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlFind.Utils
{
    public class Logger : IDisposable
    {
        /// <summary>
        /// Настроики логирования
        /// </summary>
        public LoggerSettings Settings { get; private set; }

        private FileStream _logFs;

        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="logFilename"></param>
        public Logger(string logFilename = null)
        {
            Settings = new LoggerSettings(SetLogFilename);
            Settings.LogFilename = logFilename;
            if(logFilename == null) { Settings.WriteToFile = false; }
        }

        private void SetLogFilename(string newLogFilename)
        {
            _logFs?.Close();
            if (newLogFilename != null)
            {
                _logFs = new FileStream(newLogFilename, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            }
        }
        
        /// <summary>
        /// Основная функция для записи лога
        /// </summary>
        /// <param name="msg">сообщение логирования</param>
        public void Log(string msg)
        {
            if (Settings.WriteToFile)
            {
                var logData = Encoding.UTF8.GetBytes(msg + "\r\n");
                _logFs.Write(logData, 0, logData.Length);
            }
            if (Settings.EchoConsole) { Console.WriteLine(msg); }
        }

        public void Dispose()
        {
            _logFs.Close();
        }

        public class LoggerSettings
        {
            private Action<string> _logFilenameChangedDeleg;

            public LoggerSettings(Action<string> LogFilenameChangedDeleg)
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
