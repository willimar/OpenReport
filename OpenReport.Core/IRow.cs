using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public interface IRow
    {
        public List<IField> Fields { get; }
        public int RowIndex { get; }
    }
}
