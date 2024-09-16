using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace RAG.Parsers.Docx;

internal class DocumentContext
{
    /// <summary>
    /// List of hyperlinks in the Word document
    /// </summary>
    public List<HyperlinkRelationship> Hyperlinks { get; set; } = new();

    /// <summary>
    /// List of styles referenced by a name
    /// </summary>
    public Dictionary<string, (bool isHeadingStyle, bool isTOCStyle)> DictionaryStyles { get; set; } = new();
}
