#pragma warning disable KMEXP00
using Microsoft.KernelMemory.DataFormats;
using Microsoft.KernelMemory.DataFormats.Office;
using RAG.Parsers.Xlsx;

namespace RAG.Parsers.DataIngestion.UnitTests;

/// <summary>
/// Compares the output of Microsoft.KernelMemory.Core MsExcelDecoder
/// against RAG.Parsers.Xlsx XlsxParser for the same Excel documents.
/// </summary>
[TestClass]
public class XlsxParserComparison_Should
{
    private const string TestFilesDirectory = "TestFiles";
    private const string TestFileName = "TestSample.xlsx";

    [TestMethod]
    public async Task KernelMemoryExcelDecoder_ProducesNonEmptyOutput()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);
        var decoder = new MsExcelDecoder();

        // Act
        var result = await decoder.DecodeAsync(filePath);

        // Assert
        Assert.IsNotNull(result, "KernelMemory MsExcelDecoder result should not be null.");
        Assert.IsTrue(result.Sections.Count > 0,
            "KernelMemory MsExcelDecoder should produce at least one content section.");
        Assert.IsTrue(result.Sections.Any(s => !string.IsNullOrWhiteSpace(s.Content)),
            "KernelMemory MsExcelDecoder should produce at least one non-empty content section.");
    }

    [TestMethod]
    public void RagParsersXlsxParser_ProducesMarkdownOutput()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);
        using var parser = new XlsxParser();

        // Act
        var result = parser.ExcelToMarkdown(filePath);

        // Assert
        Assert.IsNotNull(result, "RAG.Parsers.Xlsx output should not be null.");
        Assert.IsFalse(string.IsNullOrWhiteSpace(result), "RAG.Parsers.Xlsx should produce non-empty output.");
    }

    [TestMethod]
    public async Task BothParsers_ProduceNonEmptyOutputForSameExcelFile()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);

        using var ragParser = new XlsxParser();
        var kmDecoder = new MsExcelDecoder();

        // Act – RAG.Parsers.Xlsx
        var ragResult = ragParser.ExcelToMarkdown(filePath);

        // Act – KernelMemory MsExcelDecoder
        var kmResult = await kmDecoder.DecodeAsync(filePath);
        var kmText = string.Join(Environment.NewLine,
            kmResult.Sections.Where(s => !s.IsSeparator).Select(s => s.Content));

        // Assert – both parsers produce non-empty content
        Assert.IsFalse(string.IsNullOrWhiteSpace(ragResult),
            "RAG.Parsers.Xlsx should produce non-empty output.");
        Assert.IsFalse(string.IsNullOrWhiteSpace(kmText),
            "KernelMemory MsExcelDecoder should produce non-empty output.");

        // Both parsers must cover the same document: they should share common words
        var ragWords = ExtractSignificantWords(ragResult);
        var kmWords = ExtractSignificantWords(kmText);
        var overlap = ragWords.Intersect(kmWords, StringComparer.OrdinalIgnoreCase).ToList();

        Assert.IsTrue(overlap.Count > 0,
            "RAG.Parsers.Xlsx and KernelMemory MsExcelDecoder outputs should share common cell values.");
    }

    [TestMethod]
    public async Task KernelMemoryExcelDecoder_OutputContainsMarkdownWorksheetHeaders()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);
        var decoder = new MsExcelDecoder();

        // Act
        var result = await decoder.DecodeAsync(filePath);
        var fullText = string.Join(Environment.NewLine,
            result.Sections.Where(s => !s.IsSeparator).Select(s => s.Content));

        // Assert – KernelMemory uses Markdown heading syntax (# Worksheet Name) for worksheet names
        Assert.IsTrue(fullText.Contains('#'),
            "KernelMemory MsExcelDecoder output should contain Markdown heading markers ('#') for worksheet names.");
    }

    [TestMethod]
    public void RagParsersXlsxParser_OutputContainsMarkdownTableSyntax()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);
        using var parser = new XlsxParser();

        // Act
        var result = parser.ExcelToMarkdown(filePath);

        // Assert – RAG.Parsers.Xlsx wraps cell values in quotes (CSV-like markdown representation)
        Assert.IsTrue(result.Contains('"'),
            "RAG.Parsers.Xlsx output should use quoted cell values in the Markdown output.");
    }

    [TestMethod]
    public async Task KernelMemoryExcelDecoder_OutputCanBeConvertedToMarkdown()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);
        var decoder = new MsExcelDecoder();

        // Act
        var result = await decoder.DecodeAsync(filePath);

        // Build a full Markdown document from KernelMemory sections
        var markdownBuilder = new System.Text.StringBuilder();
        foreach (var section in result.Sections.Where(s => !s.IsSeparator && !string.IsNullOrWhiteSpace(s.Content)))
        {
            markdownBuilder.AppendLine(section.Content);
            markdownBuilder.AppendLine();
        }
        var markdown = markdownBuilder.ToString();

        // Assert the Markdown output is well-formed and non-empty
        Assert.IsFalse(string.IsNullOrWhiteSpace(markdown),
            "Converted Markdown output from KernelMemory MsExcelDecoder should not be empty.");
    }

    /// <summary>
    /// Extracts significant words (length >= 3) from a text for overlap comparison.
    /// </summary>
    private static IEnumerable<string> ExtractSignificantWords(string text)
    {
        return text.Split(new[] { ' ', '\n', '\r', '\t', '"', ',', '#', '|' },
                StringSplitOptions.RemoveEmptyEntries)
            .Where(w => w.Length >= 3)
            .Select(w => w.Trim());
    }
}
