using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System.Text.RegularExpressions;
using ConsoleApp1.Models;
using Newtonsoft.Json;
using xlFind.Utils;
using System.Configuration;
using xlFind.Models;
using xlFind.Models.Search;
using xlFind.Models.Replace;
using xlFind.Models.Interfaces;
using SmartFormat;

namespace xlFind
{
    public class Program
    {
#if DEBUG
        const string LOG_FILENAME = "..\\..\\logs\\replaced_values ({date:HH_mm_ss dd.MM.yyyy}).txt";
        const string SEARCHING_DIRECTORY = "..\\..";
#else
        const string LOG_FILENAME = "logs\\replaced_values ({date:HH_mm_ss dd.MM.yyyy}).txt";
        const string WORKING_DIRECTORY = ".";
#endif
        const string HELP_FILEPATH = ".\\Help\\Help.txt";



        static void Main(string[] args)
        {
            var dir = GetSearchingDirectory();
            var logFileName = GetLogFilename();
            
            var help = args.Length == 0 || args[0].Equals("-h") || args[0].Equals("help") || args[0].Equals("--help");

            if (help)
            {
                Console.WriteLine(HelpMessage());
#if DEBUG
                Console.WriteLine("Нажмите на любую кнопку чтобы выйти");
                Console.ReadKey();
#endif
                return;
            }
            var searchTxt = args[0];
            var useRegexp = args.Contains("--regexp") || args.Contains("-reg");

            var replaceTagIndex = GetReplaceTagIndex(args);
            var doReplace = replaceTagIndex != -1;
            string replacement = null;
            if (doReplace)
            {
                if(args.Length < replaceTagIndex + 2)
                {
                    string msg = $"Нужно было указать значение для замены после тэга '{args[replaceTagIndex]}'";
                    Console.WriteLine(msg);
                    Console.WriteLine(HelpMessage());
#if DEBUG
                    Console.WriteLine("Нажмите на любую кнопку чтобы выйти");
                    Console.ReadKey();
#endif
                    return;
                }
                replacement = args[replaceTagIndex + 1];
            }



            
            var matches = MatchesFinder.Find(dir, searchTxt, useRegexp);
            var count = 0;
            if (!doReplace)
            {
                var logger = new Logger(null);
                logger.Settings.EchoConsole = true;
                var result = JsonConvert.SerializeObject(matches, Formatting.Indented);
                logger.Log(result);
                count = matches.SelectMany(m => m.WorksheetMatches.SelectMany(wm => wm.CellMatches)).Count();
                Console.WriteLine($"{count} совпадении!");
            }
            else
            {
                var logger = new Logger(GetLogFilename());
                logger.Settings.EchoConsole = true;
                logger.Settings.WriteToFile = true;

                var replaces = MatchesFinder.Replace(matches, searchTxt, replacement, useRegexp);
                var result = JsonConvert.SerializeObject(replaces, Formatting.Indented);
                logger.Log(result);
                count = replaces.SelectMany(m => m.WorksheetMatches.SelectMany(wm => wm.CellMatches)).Count();
                Console.WriteLine($"{count} замен!");
            }

            if (count == 0) { Console.WriteLine("Что то ничего не нашлось :-\\ Попробуй ввести по-другому"); }
#if DEBUG
            Console.WriteLine("Нажмите на любую кнопку чтобы выйти");
            Console.ReadKey();
#endif

        }

        static void ShowSearchResult(IEnumerable<FileMatch> fileMatches)
        {
            foreach (var fileMatch in fileMatches)
            {
                var indent = ">>>>";
                Console.WriteLine($"{indent} File: {fileMatch.Filepath}");
                foreach (var worksheetMatch in fileMatch.WorksheetMatches)
                {
                    indent = "    ";
                    Console.WriteLine($"{indent} Sheet: {worksheetMatch.SheetName}");
                    foreach (var cellMatch in worksheetMatch.CellMatches)
                    {
                        indent += "    ";
                        Console.WriteLine($"{indent} {cellMatch.CellAddress} => {cellMatch.Value}");
                    }
                }
            }
        }

        static int GetReplaceTagIndex(string[] args)
        {
            var replaceTagIndex = Array.IndexOf(args, "--replacewith");
            if(replaceTagIndex == -1)
            {
                replaceTagIndex = Array.IndexOf(args, "-r");
            }
            return replaceTagIndex;
        }

        static string HelpMessage()
        {
            string msg = "Необходимо ввести значения для поиска!";
            var path = GetHelpFilePath();
            if (File.Exists(path))
            {
                msg = File.ReadAllText(path);
            }
            return msg;
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
            if(filename == null) { return null; }
            var values = new { date = DateTime.Now };
            var newFilename = Smart.Format(filename, values);
            pathParts[pathParts.Length - 1] = newFilename;
            return string.Join(Path.DirectorySeparatorChar.ToString(), pathParts);
        }

        static string GetHelpFilePath()
        {
            var settings = ConfigurationManager.AppSettings;
            return settings.AllKeys.Any(s => s.Equals("helpfile")) ? settings.Get("helpfile") : HELP_FILEPATH;
        }

    }


}
