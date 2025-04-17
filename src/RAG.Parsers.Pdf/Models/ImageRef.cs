using System;
using System.Collections.Generic;
using System.Net.Http;

namespace RAG.Parsers.Pdf.Models;

public class ImageRef
{
    /// <summary>
    /// Unique identifier
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Raw content linked to the image
    /// </summary>
    public string MarkdownRaw { get; set; }
}