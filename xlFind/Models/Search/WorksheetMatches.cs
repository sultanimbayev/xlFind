using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlFind.Models.Interfaces;

namespace xlFind.Models.Search
{
    public class WorksheetMatch
    {
        public WorksheetMatch()
        {
            CellMatches = new List<ICellMatch>();
        }
        public string SheetName { get; set; }

        public List<ICellMatch> CellMatches { get; set; }

        public WorksheetMatch Clone()
        {
            return new WorksheetMatch()
            {
                SheetName = SheetName,
                CellMatches = CellMatches.Select(cm => cm.Clone()).ToList()
            };
        }
    }
}
