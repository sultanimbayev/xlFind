using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlFind.Models.Search
{
    public class FileMatch
    {

        public FileMatch()
        {
            WorksheetMatches = new List<WorksheetMatch>();
        }
        public string Filepath { get; set; }

        public List<WorksheetMatch> WorksheetMatches { get; set; }

        public FileMatch Clone()
        {
            return new FileMatch()
            {
                Filepath = Filepath,
                WorksheetMatches = WorksheetMatches.Select(wm => wm.Clone()).ToList()
            };
        }
    }
}
