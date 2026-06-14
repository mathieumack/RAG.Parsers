using System;
using System.Collections.Generic;

namespace RAG.Parsers.Rtf.Models;

/// <summary>
/// Reference to an image extracted from an RTF document.
/// </summary>
public class ImageRef
{
    /// <summary>
    /// Unique identifier
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Image format (e.g. png, jpeg, ...)
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// Raw markdown content linked to the image
    /// </summary>
    public string MarkdownRaw { get; set; }

    /// <summary>
    /// Raw bytes for image
    /// </summary>
    public IReadOnlyList<byte> RawBytes { get; set; }
}
