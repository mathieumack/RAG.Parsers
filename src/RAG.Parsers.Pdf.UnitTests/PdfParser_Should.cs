using Microsoft.Extensions.Logging;
using RAG.Parsers.Pdf.Models;

namespace RAG.Parsers.Pdf.UnitTests;

[TestClass]
public class PdfParser_Should
{
    [TestMethod]
    [DataRow("TestFiles/TestSample.pdf", DisplayName = "Sample document")]
    [DataRow("TestFiles/TestSample2Columns.pdf", DisplayName = "Sample document on 2 columns")]
    public void ConvertPdfToMarkdown_WhenOK(string documentPath)
    {
        // Arrange
        using var loggerFactory = LoggerFactory.Create(loggingBuilder => loggingBuilder
                    .SetMinimumLevel(LogLevel.Trace)
                    .AddConsole());
        var parser = new PdfParser(loggerFactory.CreateLogger<PdfParser>());
        var filePath = Path.Combine(Environment.CurrentDirectory, documentPath);

        var result2 = File.Exists(filePath);

        // Act
        var result = parser.ToMarkdown(filePath, new ExtractOptions()
        {
            ExtractTables = true,
            ExtractPageImages = true,
            ExtractImages = true
        });

        // Assert
        Assert.IsNotNull(result.Output);
    }
}