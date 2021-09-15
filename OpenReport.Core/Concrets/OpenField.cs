using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenReport.Core.Concrets
{
    public class OpenField : IField
    {
        private readonly string _variableFormat;

        public OpenField(string cellAddress, string fieldValue)
        {
            this.CellAddress = cellAddress;
            this.OriginalValue = fieldValue;

            this._variableFormat = @"\${([a-zA-Z._]+)}";

            var addressInfo = this.GetAddressInfo(cellAddress, fieldValue);

            this.RowPos = addressInfo.rowPos;
            this.ColPos = addressInfo.colPos;
            this.IsVariable = addressInfo.isVariable;

            if (this.IsVariable)
            {
                this.FieldName = this.GetFieldName(fieldValue);
            }
            else
            {
                this.FieldName = $"field{cellAddress}";
            }
        }

        public string CellAddress { get; private set; }
        public uint ColPos { get; private set; }
        public uint RowPos { get; private set; }
        public string FieldName { get; private set; }
        public string OriginalValue { get; private set; }
        public object Formula { get; set; }
        public IStyle Style { get; set; }
        public bool IsVariable { get; private set; }

        public (uint colPos, uint rowPos, bool isVariable) GetAddressInfo(string address, string value)
        {
            var regex = new Regex("([a-zA-Z]+)([0-9]+)");
            var match = regex.Match(address);

            if (!match.Success || match.Groups.Count != 3)
            {
                throw new ArgumentException($"Invalid address to {address}.");
            }

            uint colPos = OpenField.GetColPos(match.Groups[1].Value);
            uint rowPos = Convert.ToUInt32(match.Groups[2].Value);
            bool isVariable = IsVariableCheck(value);
            return (colPos, rowPos, isVariable);
        }

        private bool IsVariableCheck(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            var regex = new Regex(this._variableFormat);
            var match = regex.Match(value);

            return match.Success;
        }

        public static uint GetColPos(string address)
        {
            if (string.IsNullOrEmpty(address))
            {
                throw new ArgumentNullException("columnName");
            }

            address = address.ToUpperInvariant();
            var regex = new Regex("([A-Z]+)");
            var match = regex.Match(address);

            address = match.Groups[1].Value;

            uint sum = 0;

            for (int i = 0; i < address.Length; i++)
            {
                sum *= 26;
                sum += (uint)address[i] - 'A' + 1;
            }

            return sum;
        }

        private string GetFieldName(string value)
        {
            if (string.IsNullOrEmpty(this._variableFormat) || string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var regex = new Regex(this._variableFormat);
            var match = regex.Match(value);

            if (!match.Success || match.Groups.Count != 2)
            {
                return string.Empty;
            }

            return match.Groups[1].Value;
        }
    }

    public class VariableField : OpenField, IVariableField
    {
        public VariableField(IField field): base(field.CellAddress, field.OriginalValue)
        {
            this.Formula = field.Formula;
            this.Style = field.Style;

            Values = new List<FieldValues>();
        }

        public ICollection<FieldValues> Values { get; private set; }
    }
}
