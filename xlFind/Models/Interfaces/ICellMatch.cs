using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlFind.Models.Interfaces
{
    public interface ICellMatch
    { 
        string CellAddress { get; set; }
        string Value { get; set; }
        ICellMatch Clone();
    }
}
