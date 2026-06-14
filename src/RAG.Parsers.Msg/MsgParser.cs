using Microsoft.Extensions.Logging;
using MsgReader.Mime.Header;
using MsgReader.Outlook;
using RAG.Parsers.Msg.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RAG.Parsers.Msg;

/// <summary>
/// Parser for Outlook .msg email files.
/// Extracts headers, body, and attachment references from the message.
/// </summary>
public class MsgParser : IDisposable
{
    private readonly ILogger<MsgParser> logger;

    public MsgParser(ILogger<MsgParser> logger)
    {
        this.logger = logger;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    #region Public Methods

    /// <summary>
    /// Parses a .msg file from the given file path.
    /// </summary>
    /// <param name="filePath">Absolute path to the .msg file.</param>
    /// <param name="options">Extraction options.</param>
    /// <returns>Structured output containing headers, body, and attachment references.</returns>
    public ExtractOutput ParseMsg(string filePath, ExtractOptions options)
    {
        using var stream = File.OpenRead(filePath);
        return ParseMsg(stream, options);
    }

    /// <summary>
    /// Parses a .msg file from the given stream.
    /// </summary>
    /// <param name="data">Stream containing the .msg file bytes.</param>
    /// <param name="options">Extraction options.</param>
    /// <returns>Structured output containing headers, body, and attachment references.</returns>
    public ExtractOutput ParseMsg(Stream data, ExtractOptions options)
    {
        var result = new ExtractOutput();

        try
        {
            using var message = new Storage.Message(data);

            ExtractSubjectAndDate(message, result);
            ExtractHeaders(message, result);
            ExtractBody(message, result);

            if (options.ExtractAttachments)
                ExtractAttachments(message, result);

            result.Output = BuildOutput(result);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to parse .msg file.");
            throw;
        }

        return result;
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
    }

    #endregion

    #region Private Methods

    private static void ExtractSubjectAndDate(Storage.Message message, ExtractOutput result)
    {
        result.Subject = message.Subject ?? "";
        result.Date = (message.SentOn ?? message.ReceivedOn)?.UtcDateTime;
        result.MessageId = message.Headers?.MessageId;
    }

    private static void ExtractHeaders(Storage.Message message, ExtractOutput result)
    {
        var headers = new EmailHeaders();

        // Sender
        if (message.Sender != null)
            headers.From = FormatSender(message.Sender);

        // Recipients
        foreach (var recipient in message.Recipients ?? [])
        {
            var address = FormatRecipient(recipient);
            if (string.IsNullOrWhiteSpace(address))
                continue;

            switch (recipient.Type)
            {
                case RecipientType.To:
                    headers.To.Add(address);
                    break;
                case RecipientType.Cc:
                    headers.Cc.Add(address);
                    break;
                case RecipientType.Bcc:
                    headers.Bcc.Add(address);
                    break;
            }
        }

        // RFC 822 / MIME headers (available when the message was originally an internet email)
        if (message.Headers != null)
        {
            if (message.Headers.ReplyTo is { HasValidMailAddress: true })
                headers.ReplyTo = message.Headers.ReplyTo.Address;

            headers.ContentType = message.Headers.ContentType?.ToString();
            headers.ContentTransferEncoding = message.Headers.ContentTransferEncoding.ToString();

            // Fall back to RFC headers for To/Cc/Bcc if MAPI recipients list is empty
            if (headers.To.Count == 0)
                headers.To.AddRange(message.Headers.To.Select(r => r.HasValidMailAddress ? r.Address : r.DisplayName).Where(s => !string.IsNullOrWhiteSpace(s)));
            if (headers.Cc.Count == 0)
                headers.Cc.AddRange(message.Headers.Cc.Select(r => r.HasValidMailAddress ? r.Address : r.DisplayName).Where(s => !string.IsNullOrWhiteSpace(s)));
            if (headers.Bcc.Count == 0)
                headers.Bcc.AddRange(message.Headers.Bcc.Select(r => r.HasValidMailAddress ? r.Address : r.DisplayName).Where(s => !string.IsNullOrWhiteSpace(s)));
        }

        result.Headers = headers;
    }

    private static void ExtractBody(Storage.Message message, ExtractOutput result)
    {
        result.Body = new EmailBody
        {
            Text = string.IsNullOrWhiteSpace(message.BodyText) ? null : message.BodyText,
            Html = string.IsNullOrWhiteSpace(message.BodyHtml) ? null : message.BodyHtml
        };
    }

    private static void ExtractAttachments(Storage.Message message, ExtractOutput result)
    {
        foreach (var item in message.Attachments ?? [])
        {
            if (item is not Storage.Attachment attachment)
                continue;

            // Skip hidden/inline attachments that are purely decorative
            if (attachment.Hidden)
                continue;

            var attachRef = new AttachmentRef
            {
                FileName = attachment.FileName ?? "",
                ContentType = attachment.MimeType ?? "application/octet-stream",
                RawBytes = attachment.Data ?? [],
                Size = attachment.Data?.Length ?? 0
            };

            result.Attachments.Add(attachRef);
        }
    }

    /// <summary>
    /// Builds the flat Markdown-style text output used by downstream RAG pipelines.
    /// </summary>
    private static string BuildOutput(ExtractOutput result)
    {
        var sb = new StringBuilder();

        sb.AppendLine($"# {result.Subject}");
        sb.AppendLine();

        if (!string.IsNullOrWhiteSpace(result.Headers.From))
            sb.AppendLine($"**From:** {result.Headers.From}");

        if (result.Headers.To.Count > 0)
            sb.AppendLine($"**To:** {string.Join("; ", result.Headers.To)}");

        if (result.Headers.Cc.Count > 0)
            sb.AppendLine($"**Cc:** {string.Join("; ", result.Headers.Cc)}");

        if (result.Headers.Bcc.Count > 0)
            sb.AppendLine($"**Bcc:** {string.Join("; ", result.Headers.Bcc)}");

        if (result.Date.HasValue)
            sb.AppendLine($"**Date:** {result.Date.Value:O}");

        if (!string.IsNullOrWhiteSpace(result.MessageId))
            sb.AppendLine($"**Message-ID:** {result.MessageId}");

        if (result.Attachments.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("**Attachments:**");
            foreach (var att in result.Attachments)
                sb.AppendLine($"- {att.FileName} ({att.ContentType}, {att.Size} bytes)");
        }

        sb.AppendLine();

        // Prefer plain text; fall back to HTML when text is absent
        var bodyText = result.Body.Text;
        if (!string.IsNullOrWhiteSpace(bodyText))
            sb.Append(bodyText);
        else if (!string.IsNullOrWhiteSpace(result.Body.Html))
            sb.Append(result.Body.Html);

        return sb.ToString().Trim();
    }

    private static string FormatSender(Storage.Sender sender)
    {
        if (!string.IsNullOrWhiteSpace(sender.DisplayName) && !string.IsNullOrWhiteSpace(sender.Email))
            return $"{sender.DisplayName} <{sender.Email}>";
        if (!string.IsNullOrWhiteSpace(sender.Email))
            return sender.Email;
        return sender.DisplayName ?? "";
    }

    private static string FormatRecipient(Storage.Recipient recipient)
    {
        if (!string.IsNullOrWhiteSpace(recipient.DisplayName) && !string.IsNullOrWhiteSpace(recipient.Email))
            return $"{recipient.DisplayName} <{recipient.Email}>";
        if (!string.IsNullOrWhiteSpace(recipient.Email))
            return recipient.Email;
        return recipient.DisplayName ?? "";
    }

    #endregion
}
