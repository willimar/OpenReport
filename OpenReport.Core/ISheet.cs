using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public interface ISheet
    {
        string Name { get; set; }

        IEnumerable<IField> Fields { get; }
    }
}
