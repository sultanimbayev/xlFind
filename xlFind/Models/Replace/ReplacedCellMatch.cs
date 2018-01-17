using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlFind.Models.Interfaces;

namespace xlFind.Models.Replace
{
    public class ReplacedCellMatch : ICellMatch
    {
        public string CellAddress { get; set; }
        public string Value { get; set; }
        public string OldValue { get; set; }
        public ICellMatch Clone()
        {
            return new ReplacedCellMatch()
            {
                CellAddress = CellAddress,
                Value = Value,
                OldValue = Value
            };
        }
    }
}
