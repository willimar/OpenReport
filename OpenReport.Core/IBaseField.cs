namespace OpenReport.Core
{
    public interface IBaseField
    {
        string CellAddress { get; }
        uint ColPos { get; }
        uint RowPos { get; }
        string FieldName { get; }
        string OriginalValue { get; }
        object Formula { get; set; }
        IStyle Style { get; set; }
        bool IsVariable { get; }
    }
}