using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core.Concrets
{
    public class OpenRow : IRow
    {
        public OpenRow(int rowIndex)
        {
            this.Fields = new List<IField>();
            this.RowIndex = rowIndex;
        }

        public List<IField> Fields { get; }
        public int RowIndex { get; private set; }
    }
}
