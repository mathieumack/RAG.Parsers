#pragma warning disable KMEXP00
using Microsoft.Extensions.Logging;
using Microsoft.KernelMemory.DataFormats;
using Microsoft.KernelMemory.DataFormats.Pdf;
using RAG.Parsers.Pdf;
using RAG.Parsers.Pdf.Models;

namespace RAG.Parsers.DataIngestion.UnitTests;

/// <summary>
/// Compares the output of Microsoft.KernelMemory.Core PdfDecoder
/// against RAG.Parsers.Pdf PdfParser for the same PDF documents.
/// </summary>
[TestClass]
public class PdfParserComparison_Should
{
    private const string TestFilesDirectory = "TestFiles";

    [TestMethod]
    [DataRow("TestSample.pdf", DisplayName = "Single-column PDF document")]
    [DataRow("TestSample2Columns.pdf", DisplayName = "Two-column PDF document")]
    public async Task KernelMemoryPdfDecoder_ProducesNonEmptyOutput(string fileName)
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, fileName);
        var decoder = new PdfDecoder();

        // Act
        var result = await decoder.DecodeAsync(filePath);

        // Assert
        Assert.IsNotNull(result, "KernelMemory PdfDecoder result should not be null.");
        Assert.IsTrue(result.Sections.Count > 0, "KernelMemory PdfDecoder should produce at least one content section.");
        Assert.IsTrue(result.Sections.Any(s => !string.IsNullOrWhiteSpace(s.Content)),
            "KernelMemory PdfDecoder should produce at least one non-empty content section.");
    }

    [TestMethod]
    [DataRow("TestSample.pdf", DisplayName = "Single-column PDF document")]
    [DataRow("TestSample2Columns.pdf", DisplayName = "Two-column PDF document")]
    public void RagParsersPdfParser_ProducesMarkdownOutput(string fileName)
    {
        // Arrange
        using var loggerFactory = LoggerFactory.Create(builder =>
            builder.SetMinimumLevel(LogLevel.Warning).AddConsole());
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, fileName);
        using var parser = new PdfParser(loggerFactory.CreateLogger<PdfParser>());

        // Act
        var result = parser.ToMarkdown(filePath, new ExtractOptions
        {
            ExtractTables = true,
            ExtractImages = false,
            ExtractPageImages = false
        });

        // Assert
        Assert.IsNotNull(result, "RAG.Parsers.Pdf output should not be null.");
        Assert.IsNotNull(result.Output, "RAG.Parsers.Pdf output text should not be null.");
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output), "RAG.Parsers.Pdf should produce non-empty output.");
    }

    [TestMethod]
    [DataRow("TestSample.pdf", DisplayName = "Single-column PDF document")]
    [DataRow("TestSample2Columns.pdf", DisplayName = "Two-column PDF document")]
    public async Task BothParsers_ProduceNonEmptyOutputForSamePdfFile(string fileName)
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, fileName);

        using var loggerFactory = LoggerFactory.Create(builder =>
            builder.SetMinimumLevel(LogLevel.Warning).AddConsole());
        using var ragParser = new PdfParser(loggerFactory.CreateLogger<PdfParser>());
        var kmDecoder = new PdfDecoder();

        // Act – RAG.Parsers.Pdf
        var ragResult = ragParser.ToMarkdown(filePath, new ExtractOptions
        {
            ExtractTables = true,
            ExtractImages = false,
            ExtractPageImages = false
        });

        // Act – KernelMemory PdfDecoder
        var kmResult = await kmDecoder.DecodeAsync(filePath);
        var kmText = string.Join(Environment.NewLine,
            kmResult.Sections.Where(s => !s.IsSeparator).Select(s => s.Content));

        // Assert – both parsers produce non-empty content
        Assert.IsFalse(string.IsNullOrWhiteSpace(ragResult.Output),
            "RAG.Parsers.Pdf should produce non-empty output.");
        Assert.IsFalse(string.IsNullOrWhiteSpace(kmText),
            "KernelMemory PdfDecoder should produce non-empty output.");

        // Both parsers must cover the same document: they should share common words
        var ragWords = ExtractSignificantWords(ragResult.Output);
        var kmWords = ExtractSignificantWords(kmText);
        var overlap = ragWords.Intersect(kmWords, StringComparer.OrdinalIgnoreCase).ToList();

        Assert.IsTrue(overlap.Count > 0,
            $"RAG.Parsers.Pdf and KernelMemory PdfDecoder outputs should share common words for '{fileName}'.");
    }

    [TestMethod]
    [DataRow("TestSample.pdf", DisplayName = "Single-column PDF document")]
    [DataRow("TestSample2Columns.pdf", DisplayName = "Two-column PDF document")]
    public async Task KernelMemoryPdfDecoder_OutputCanBeConvertedToMarkdown(string fileName)
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, fileName);
        var decoder = new PdfDecoder();

        // Act
        var result = await decoder.DecodeAsync(filePath);

        // Convert KernelMemory sections to a simple Markdown document
        var markdownBuilder = new System.Text.StringBuilder();
        foreach (var section in result.Sections.Where(s => !s.IsSeparator && !string.IsNullOrWhiteSpace(s.Content)))
        {
            if (section.PageNumber > 0)
                markdownBuilder.AppendLine($"## Page {section.PageNumber}");

            markdownBuilder.AppendLine(section.Content);
            markdownBuilder.AppendLine();
        }
        var markdown = markdownBuilder.ToString();

        // Assert the markdown output is well-formed
        Assert.IsFalse(string.IsNullOrWhiteSpace(markdown),
            "Converted Markdown output from KernelMemory PdfDecoder should not be empty.");
    }

    /// <summary>
    /// Extracts significant words (length >= 4) from a text for overlap comparison.
    /// </summary>
    private static IEnumerable<string> ExtractSignificantWords(string text)
    {
        return text.Split(new[] { ' ', '\n', '\r', '\t', '.', ',', ';', ':' },
                StringSplitOptions.RemoveEmptyEntries)
            .Where(w => w.Length >= 4)
            .Select(w => w.Trim());
    }
}
