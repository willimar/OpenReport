using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public interface IRepeatGroup
    {
        uint StartRow { get; set; }
        uint EndRow { get; set; }

        IEnumerable<IEnumerable<IBaseField>> Groups { get; set; }
    }
}
