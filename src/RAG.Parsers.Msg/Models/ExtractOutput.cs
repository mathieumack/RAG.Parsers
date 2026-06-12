using System;
using System.Collections.Generic;

namespace RAG.Parsers.Msg.Models;

/// <summary>
/// Output produced by <see cref="RAG.Parsers.Msg.MsgParser"/>.
/// </summary>
public class ExtractOutput
{
    /// <summary>
    /// Structured email headers (From, To, Cc, Bcc, Reply-To, …).
    /// </summary>
    public EmailHeaders Headers { get; set; } = new();

    /// <summary>
    /// Email subject line.
    /// </summary>
    public string Subject { get; set; } = "";

    /// <summary>
    /// Sent date/time of the message, when available.
    /// </summary>
    public DateTime? Date { get; set; }

    /// <summary>
    /// Message-ID header value, when available.
    /// </summary>
    public string? MessageId { get; set; }

    /// <summary>
    /// Message body in plain-text and/or HTML form.
    /// </summary>
    public EmailBody Body { get; set; } = new();

    /// <summary>
    /// References to attachments extracted from the message.
    /// Raw bytes are included so the caller can store/upload them.
    /// The caller should populate <see cref="AttachmentRef.ImageRef"/> after storing.
    /// </summary>
    public List<AttachmentRef> Attachments { get; set; } = [];

    /// <summary>
    /// Flat text output suitable for downstream RAG / embedding pipelines.
    /// Contains subject, headers, and plain-text body concatenated in a
    /// human-readable Markdown-style format.
    /// </summary>
    public string Output { get; set; } = "";
}
