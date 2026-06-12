namespace RAG.Parsers.Msg.Models;

/// <summary>
/// Contains the plain-text and HTML representations of the message body.
/// </summary>
public class EmailBody
{
    /// <summary>
    /// Plain-text body. Preferred for text-only consumers.
    /// </summary>
    public string? Text { get; set; }

    /// <summary>
    /// HTML body, when available.
    /// </summary>
    public string? Html { get; set; }
}
