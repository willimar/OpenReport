namespace OpenReport.Core
{
    public interface IBaseField
    {
        string CellAddress { get; }
        uint ColPos { get; }
        uint RowPos { get; }
        string FieldName { get; }
        string FieldValue { get; set; }
        object Formula { get; set; }
        IStyle Style { get; set; }
        bool IsVariable { get; }
    }
}