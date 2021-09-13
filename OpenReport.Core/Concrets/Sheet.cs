using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core.Concrets
{
    internal class Sheet : ISheet
    {
        public Sheet()
        {

        }

        public string Name { get; set; }
        public IEnumerable<IField> Fields { get; set; }
    }
}
