using System.Collections.Generic;

namespace RAG.Parsers.Msg.Models;

/// <summary>
/// Represents attachment metadata extracted from a .msg file.
/// The caller is responsible for storing/reusing the raw bytes as images
/// or other artifacts and for passing references between parsers.
/// </summary>
public class AttachmentRef
{
    /// <summary>
    /// Original filename of the attachment.
    /// </summary>
    public string FileName { get; set; } = "";

    /// <summary>
    /// MIME content type of the attachment (e.g. "application/pdf", "image/png").
    /// </summary>
    public string ContentType { get; set; } = "";

    /// <summary>
    /// Size of the attachment in bytes.
    /// </summary>
    public long Size { get; set; }

    /// <summary>
    /// Raw bytes of the attachment.
    /// Exposed so the caller can upload/store the attachment and reuse it
    /// (e.g. pass to another parser or store as an image reference).
    /// </summary>
    public IReadOnlyList<byte> RawBytes { get; set; } = [];

    /// <summary>
    /// Caller-managed image/document reference string.
    /// Populate this after uploading or storing the attachment elsewhere.
    /// </summary>
    public string? ImageRef { get; set; }
}
