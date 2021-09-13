using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public interface IBookProvider
    {
        string[] GetSheetList();
        IEnumerable<(string cellAddress, string variableName)> GetVariableList(string sheetName);
        ValueTask<object> GetFieldValueTo(string cellAddress, string variableName, string sheetName);
        ValueTask<IStyle> GetFieldStyleTo(string cellAddress, string variableName, string sheetName);
        ValueTask<object> GetFormulaTo(string cellAddress, string variableName, string sheetName);
        ValueTask<ISheetBuilder> SheetBuilderCreate(string name);
        ValueTask<FieldType> GetFieldType(string cellAddress, string variableName, string sheetName);
        ValueTask<byte[]> GetBuildedReport();
    }
}
