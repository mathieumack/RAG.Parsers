using System.Collections.Generic;

namespace RAG.Parsers.Rtf.Models;

/// <summary>
/// Output document that contains all extracted information
/// </summary>
public class ExtractOutput
{
    /// <summary>
    /// List of refs to images extracted from the document
    /// </summary>
    public List<ImageRef> Images { get; set; } = new();

    /// <summary>
    /// Output text in markdown format
    /// </summary>
    public string Output { get; set; } = "";
}
