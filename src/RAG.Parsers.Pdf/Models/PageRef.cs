using System;
using System.Collections.Generic;

namespace RAG.Parsers.Pdf.Models;

public class PageRef
{
    /// <summary>
    /// Unique identifier
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Page number, starts at 1
    /// </summary>
    public int PageNumber { get; set; }

    /// <summary>
    /// Image format (e,g. png, jpeg, ...)
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// Raw content linked to the image that represents the page
    /// </summary>
    public string MarkdownRaw { get; set; }

    /// <summary>
    /// Raw bytes for image
    /// </summary>
    public IReadOnlyList<Byte> RawBytes { get; set; }
}
