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
    }

    public class FieldValues
    {
        public FieldValues(int order, string value)
        {
            this.Order = order;
            this.Value = value;
        }

        public int Order { get; }
        public string Value { get; }
    }

    public interface IVariableField : IField
    {
        ICollection<FieldValues> Values { get; }
    }
}
