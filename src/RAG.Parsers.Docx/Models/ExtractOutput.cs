using DocumentFormat.OpenXml.Packaging;
using RAG.Parsers.Docx.Models;
using System.Collections.Generic;

namespace RAG.Parsers.Docx;

/// <summary>
/// Output document that contains all extracted informations
/// </summary>
public class ExtractOutput
{
    /// <summary>
    /// List of hyperlinks in the Word document
    /// </summary>
    internal List<HyperlinkRelationship> Hyperlinks { get; set; } = new();

    /// <summary>
    /// List of refs to image
    /// </summary>
    public List<ImageRef> Images { get; set; } = new();

    /// <summary>
    /// List of styles referenced by a name
    /// </summary>
    internal Dictionary<string, (bool isHeadingStyle, bool isTOCStyle)> DictionaryStyles { get; set; } = new();

    /// <summary>
    /// Output text
    /// </summary>
    public string Output { get; set; } = "";
}
