using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core.Concrets
{
    public class OpenSheet : ISheet
    {
        public OpenSheet(string name)
        {
            this.Name = name;
            this.Rows = new List<IRow>();
        }

        public string Name { get; }
        public List<IRow> Rows { get; }
    }
}
