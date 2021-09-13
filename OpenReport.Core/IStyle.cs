using System.Drawing;

namespace OpenReport.Core
{
    public interface IStyle
    {
        uint StyleIndex { get; set; }
        Font Font { get; set; }
        string MergeCells { get; set; }
    }
}