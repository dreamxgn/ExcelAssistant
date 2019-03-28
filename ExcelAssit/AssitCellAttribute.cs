using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAssit
{
    public class AssitCellAttribute:Attribute
    {
        public string CellTitle { get; set; }
        public CellType CellType { get; set; }
    }

    public enum CellType
    {
        String,
        Int,
        Float,
        Image,
        DateTime
    }


}
