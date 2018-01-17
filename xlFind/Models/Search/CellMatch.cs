using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlFind.Models.Interfaces;

namespace xlFind.Models.Search
{
    public class CellMatch : ICellMatch
    {
        public string CellAddress { get; set; }
        public string Value { get; set; }

        public ICellMatch Clone()
        {
            return new CellMatch()
            {
                CellAddress = CellAddress,
                Value = Value
            };
        }
    }
}
