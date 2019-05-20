using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlFind.FindText
{
    [Verb("find-text")]
    public class FindTextOptions
    {
        [Value(0, HelpText = "Text to find in excel files", Required = true )]
        public string SearchText { get; set; }

        [Option('r', "replace-with", Default = null, HelpText = "Text to replace with")]
        public string ReplaceText { get; set; }

        [Option("rgx", Default = false, HelpText = "If set to 'true', then regular expressions will be used")]
        public bool UseRegex { get; set; }
    }
}
