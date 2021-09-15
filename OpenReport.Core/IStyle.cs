using System.Drawing;

namespace OpenReport.Core
{
    public interface IStyle
    {
        int StyleIndex { get; set; }
        bool IsMerged { get; set; }
        string MergeCells { get; set; }
    }
}