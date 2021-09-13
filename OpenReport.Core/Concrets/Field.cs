using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenReport.Core.Concrets
{
    internal class Field : IField
    {
        public Field(string cellAddress, string variableName)
        {
            this.CellAddress = cellAddress;
            this.VariableName = variableName;
            
            (uint colPos, uint rowPos) addressInfo = Field.GetAddressInfo(cellAddress);

            this.RowPos = addressInfo.rowPos;
            this.ColPos = addressInfo.colPos;
            this.VariableFormat = @"\${([a-zA-Z._]+)}";
        }

        public string CellAddress { get; private set; }
        public uint ColPos { get; private set; }
        public uint RowPos { get; private set; }
        public string VariableName { get; private set; }
        public string VariableFormat { get; set; }
        public string FieldName { get { return this.GetFieldName(); } }
        public object FieldValue { get; set; }
        public object Formula { get; set; }
        public IStyle Style { get; set; }
        public FieldType FieldType { get; set; }

        public static (uint colPos, uint rowPos) GetAddressInfo(string address)
        {
            var regex = new Regex("([a-zA-Z]+)([0-9]+)");
            var match = regex.Match(address);

            if (!match.Success || match.Groups.Count != 3)
            {
                throw new ArgumentException($"Invalid address to {address}.");
            }

            uint colPos = Field.GetColPos(match.Groups[1].Value);
            uint rowPos = Convert.ToUInt32(match.Groups[2].Value);
            return (colPos, rowPos);
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

        private string GetFieldName()
        {
            if (string.IsNullOrEmpty(this.VariableFormat) || string.IsNullOrEmpty(this.VariableName))
            {
                return string.Empty;
            }

            var regex = new Regex(this.VariableFormat);
            var match = regex.Match(this.VariableName);

            if (!match.Success || match.Groups.Count != 2)
            {
                return string.Empty;
            }

            return match.Groups[1].Value;
        }
    }
}
