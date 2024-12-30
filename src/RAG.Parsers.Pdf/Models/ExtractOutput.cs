using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAG.Parsers.Pdf.Models;

public class ExtractOutput
{
    /// <summary>
    /// List of refs to image
    /// </summary>
    public List<ImageRef> Images { get; set; } = new();

    /// <summary>
    /// Output text
    /// </summary>
    public string Output { get; set; } = "";
}
