using Html2Markdown;
using RAG.Parsers.Rtf.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace RAG.Parsers.Rtf;

/// <summary>
/// RTF Decoder to Markdown
/// </summary>
public class RtfParser : IDisposable
{
    static RtfParser()
    {
        // RtfPipe requires Windows code page encodings (e.g. Windows-1252)
        // which are not available in .NET Core by default.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    #region Public Methods

    /// <summary>
    /// Read file and convert it to markdown
    /// </summary>
    /// <param name="filePath">Path to the .rtf file</param>
    /// <param name="options">Extraction options</param>
    /// <returns>Extracted content as markdown text</returns>
    public ExtractOutput ToMarkdown(string filePath, ExtractOptions options = null)
    {
        using var stream = File.OpenRead(filePath);
        return ToMarkdown(stream, options);
    }

    /// <summary>
    /// Convert RTF stream to markdown
    /// </summary>
    /// <param name="data">RTF data stream</param>
    /// <param name="options">Extraction options</param>
    /// <returns>Extracted content as markdown text</returns>
    public ExtractOutput ToMarkdown(Stream data, ExtractOptions options = null)
    {
        using var reader = new StreamReader(data, Encoding.UTF8, true, 1024, leaveOpen: true);
        var rtfContent = reader.ReadToEnd();
        return ConvertRtfContent(rtfContent);
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Converts RTF string content to an ExtractOutput with markdown text and extracted images.
    /// </summary>
    private static ExtractOutput ConvertRtfContent(string rtfContent)
    {
        var context = new ExtractOutput();

        if (string.IsNullOrWhiteSpace(rtfContent))
        {
            return context;
        }

        try
        {
            var html = RtfPipe.Rtf.ToHtml(rtfContent);
            var markdown = HtmlToMarkdown(html, context.Images);
            context.Output = markdown.Trim();
        }
        catch (Exception)
        {
            context.Output = string.Empty;
        }

        return context;
    }

    /// <summary>
    /// Converts HTML output from RtfPipe into markdown, extracting embedded images.
    /// </summary>
    private static string HtmlToMarkdown(string html, List<ImageRef> images)
    {
        if (string.IsNullOrWhiteSpace(html))
            return string.Empty;

        // Strip the outer <div> wrapper that RtfPipe wraps its output in,
        // since Html2Markdown cannot map it to any markdown construct.
        html = StripOuterDiv(html);

        // Extract embedded base64 images and replace with ID-keyed markdown references
        html = ExtractImages(html, images);

        // Convert the remaining HTML to Markdown using Html2Markdown
        var converter = new Converter();
        return converter.Convert(html);
    }

    /// <summary>
    /// Strips the outermost &lt;div&gt; element added by RtfPipe, keeping its inner HTML.
    /// </summary>
    private static string StripOuterDiv(string html)
    {
        var trimmed = html.Trim();
        var match = Regex.Match(trimmed, @"^<div[^>]*>(.*)</div>$", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : trimmed;
    }

    /// <summary>
    /// Finds &lt;img&gt; tags whose src is a base64 data URI, extracts the raw bytes into
    /// <paramref name="images"/>, and replaces the tag with a markdown image reference.
    /// </summary>
    private static string ExtractImages(string html, List<ImageRef> images)
    {
        return Regex.Replace(
            html,
            @"<img\s[^>]*src=""data:image/([^;]+);base64,([^""]+)""[^>]*/?>",
            match =>
            {
                var format = match.Groups[1].Value.ToLowerInvariant();
                var base64 = match.Groups[2].Value;

                byte[] bytes;
                try { bytes = Convert.FromBase64String(base64); }
                catch { return string.Empty; }

                var id = $"{Guid.NewGuid()}.{format}";
                var raw = $"![image](data:image/{format};{id})";

                images.Add(new ImageRef
                {
                    Id = id,
                    Format = format,
                    MarkdownRaw = raw,
                    RawBytes = Array.AsReadOnly(bytes)
                });

                return raw;
            },
            RegexOptions.IgnoreCase);
    }

    #endregion

    public void Dispose()
    {
        // Nothing to release
        GC.SuppressFinalize(this);
    }
}

