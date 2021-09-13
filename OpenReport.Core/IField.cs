using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public enum FieldType
    {
        simpleField = 1,
        repeatGroup = 2
    }

    public interface IField : IBaseField
    {
        FieldType FieldType { get; set; }
    }
}
