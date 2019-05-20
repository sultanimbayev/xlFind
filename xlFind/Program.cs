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
using CommandLine;
using xlFind.FindText;

namespace xlFind
{
    public class Program
    {
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<FindTextOptions>(args)
                .MapResult(
                    (FindTextOptions findTextOpts) => (new FindTextHandler()).Handle(findTextOpts),
                    errs => 1
                );
        }
    }


}
