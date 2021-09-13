namespace OpenReport.Core
{
    public interface IBaseField
    {
        string CellAddress { get; }
        uint ColPos { get; }
        uint RowPos { get; }
        string VariableName { get; }
        string VariableFormat { get; set; }
        string FieldName { get; }
        object FieldValue { get; set; }
        object Formula { get; set; }
        IStyle Style { get; set; }
    }
}