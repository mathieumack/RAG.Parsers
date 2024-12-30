using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAG.Parsers.Docx.Models.Table;

public class CellInfo
{
    public TableCell Cell { get; set; }

    public MergeType HorizontalMerge { get; set; }

    public MergeType VerticalMerge { get; set; }

    public int GridSpan { get; set; } = 1;

    public int VerticalSpan { get; set; } = 1;

    public bool IsOccupied { get; set; }
}
