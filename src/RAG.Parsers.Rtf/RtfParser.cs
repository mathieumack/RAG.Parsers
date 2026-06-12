using RAG.Parsers.Rtf.Models;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

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
    /// Converts RTF string content to an ExtractOutput with plain markdown text.
    /// </summary>
    private static ExtractOutput ConvertRtfContent(string rtfContent)
    {
        if (string.IsNullOrWhiteSpace(rtfContent))
        {
            return new ExtractOutput { Output = string.Empty };
        }

        try
        {
            var html = RtfPipe.Rtf.ToHtml(rtfContent);
            var plainText = HtmlToMarkdown(html);
            return new ExtractOutput { Output = plainText.Trim() };
        }
        catch (Exception)
        {
            return new ExtractOutput { Output = string.Empty };
        }
    }

    /// <summary>
    /// Converts HTML output from RtfPipe into plain markdown text,
    /// preserving paragraph and line breaks.
    /// </summary>
    private static string HtmlToMarkdown(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
            return string.Empty;

        // Replace block-level elements with newlines before stripping tags
        html = Regex.Replace(html, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"</p>", "\n\n", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"</div>", "\n", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"</h[1-6]>", "\n\n", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"</li>", "\n", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"<li[^>]*>", "- ", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"</tr>", "\n", RegexOptions.IgnoreCase);
        html = Regex.Replace(html, @"</td>|</th>", "\t", RegexOptions.IgnoreCase);

        // Strip remaining HTML tags
        html = Regex.Replace(html, @"<[^>]+>", string.Empty);

        // Decode HTML entities
        html = HttpUtility.HtmlDecode(html);

        // Normalize multiple blank lines to at most two
        html = Regex.Replace(html, @"\n{3,}", "\n\n");

        return html;
    }

    #endregion

    public void Dispose()
    {
        // Nothing to release
        GC.SuppressFinalize(this);
    }
}
