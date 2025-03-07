using System.Collections.Generic;
using System;

namespace RAG.Parsers.Docx.Models;

public class ImageRef
{
    /// <summary>
    /// Unique identifier
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Image format (e,g. png, jpeg, ...)
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// Raw content linked to the image
    /// </summary>
    public string MarkdownRaw { get; set; }

    /// <summary>
    /// Raw bytes for image
    /// </summary>
    public IReadOnlyList<Byte> RawBytes { get; set; }

    /// <summary>
    /// Image width in pixels
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// Image height in pixels
    /// </summary>
    public double Height { get; set; }

}
