using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenReport.Core.Concrets
{
    public class OpenStyle : IStyle
    {
        public int StyleIndex { get; set; }
        public string MergeCells { get; set; }
        public bool IsMerged { get; set; }
    }
}
