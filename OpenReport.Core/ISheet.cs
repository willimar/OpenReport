using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public interface ISheet
    {
        string Name { get; }

        List<IRow> Rows { get; }
        List<IVariableField> Variables { get; }
    }
}
