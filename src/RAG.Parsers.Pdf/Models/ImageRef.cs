
namespace RAG.Parsers.Pdf.Models;

public class ImageRef
{
    /// <summary>
    /// Unique identifier
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Image format name
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// Raw content linked to the image
    /// </summary>
    public string MarkdownRaw { get; set; }
}