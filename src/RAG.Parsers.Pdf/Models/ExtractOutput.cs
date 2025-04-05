using System.Collections.Generic;

namespace RAG.Parsers.Pdf.Models;

public class ExtractOutput
{
    /// <summary>
    /// List of refs to image
    /// </summary>
    public List<ImageRef> Images { get; set; } = new();

    /// <summary>
    /// List of refs to pages
    /// </summary>
    public List<PageRef> Pages { get; set; } = new();

    /// <summary>
    /// Output text
    /// </summary>
    public string Output { get; set; } = "";
}
