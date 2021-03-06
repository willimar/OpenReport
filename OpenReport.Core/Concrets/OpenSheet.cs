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
            this.Variables = new List<IVariableField>();
        }

        public string Name { get; }
        public List<IRow> Rows { get; }
        public List<IVariableField> Variables { get; }
    }
}
