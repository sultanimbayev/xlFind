using Newtonsoft.Json;
using SmartFormat;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlFind.Utils;

namespace xlFind.FindText
{
    public class FindTextHandler
    {
#if DEBUG
        const string LOG_FILENAME = "..\\..\\logs\\replaced_values ({date:HH_mm_ss dd.MM.yyyy}).txt";
        const string SEARCHING_DIRECTORY = "..\\..";
#else
        const string LOG_FILENAME = "logs\\replaced_values ({date:HH_mm_ss dd.MM.yyyy}).txt";
        const string WORKING_DIRECTORY = ".";
#endif
        const string HELP_FILEPATH = ".\\Help\\Help.txt";

        public int Handle(FindTextOptions options)
        {
            var dir = GetSearchingDirectory();
            var logFileName = GetLogFilename();
            Console.WriteLine($"looking for: \"{options.SearchText}\"");
            var matches = MatchesFinder.Find(dir, options.SearchText, options.UseRegex);
            var count = 0;
            if (string.IsNullOrEmpty(options.ReplaceText))
            {
                var logger = new Logger(null);
                logger.Settings.EchoConsole = true;
                var result = JsonConvert.SerializeObject(matches, Formatting.Indented);
                logger.Log(result);
                count = matches.SelectMany(m => m.WorksheetMatches.SelectMany(wm => wm.CellMatches)).Count();
                Console.WriteLine($"{count} matches!");
            }
            else
            {
                var logger = new Logger(GetLogFilename());
                logger.Settings.EchoConsole = true;
                logger.Settings.WriteToFile = true;

                var replaces = MatchesFinder.Replace(matches, options.SearchText, options.ReplaceText, options.UseRegex);
                var result = JsonConvert.SerializeObject(replaces, Formatting.Indented);
                logger.Log(result);
                count = replaces.SelectMany(m => m.WorksheetMatches.SelectMany(wm => wm.CellMatches)).Count();
                Console.WriteLine($"{count} replaces!");
            }

            if (count == 0) { Console.WriteLine("Nothing found :-\\ try another word :))"); }
            Console.WriteLine("Press any key to quit...");
            Console.ReadKey();

            return 0;
        }


        static string GetSearchingDirectory()
        {
            var settings = ConfigurationManager.AppSettings;
            return settings.AllKeys.Any(s => s.Equals("dir")) ? settings.Get("dir") : ".";
        }


        static string GetLogFilename()
        {
            var settings = ConfigurationManager.AppSettings;
            string format = settings.AllKeys.Any(s => s.Equals("logfile")) ? settings.Get("logfile") : LOG_FILENAME;
            var pathParts = format.Split(Path.DirectorySeparatorChar);
            var filename = pathParts.LastOrDefault();
            if (filename == null) { return null; }
            var values = new { date = DateTime.Now };
            var newFilename = Smart.Format(filename, values);
            pathParts[pathParts.Length - 1] = newFilename;
            return string.Join(Path.DirectorySeparatorChar.ToString(), pathParts);
        }


    }
}
