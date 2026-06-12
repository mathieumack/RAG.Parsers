using System.Collections.Generic;

namespace RAG.Parsers.Msg.Models;

/// <summary>
/// Contains the extracted email headers.
/// </summary>
public class EmailHeaders
{
    /// <summary>Sender address.</summary>
    public string From { get; set; } = "";

    /// <summary>Primary recipients.</summary>
    public List<string> To { get; set; } = [];

    /// <summary>Carbon-copy recipients.</summary>
    public List<string> Cc { get; set; } = [];

    /// <summary>Blind carbon-copy recipients.</summary>
    public List<string> Bcc { get; set; } = [];

    /// <summary>Reply-To address, when present.</summary>
    public string? ReplyTo { get; set; }

    /// <summary>Raw Content-Type header value, when present.</summary>
    public string? ContentType { get; set; }

    /// <summary>Raw Content-Transfer-Encoding header value, when present.</summary>
    public string? ContentTransferEncoding { get; set; }
}
