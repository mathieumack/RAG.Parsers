using RAG.Parsers.Pdf.Models;

namespace RAG.Parsers.Pdf.UnitTests;

[TestClass]
public class PdfParser_Should
{
    [TestMethod]
    [DataRow("TestFiles/TestSample.pdf", DisplayName = "Sample document")]
    [DataRow("TestFiles/TestSample2Columns.pdf", DisplayName = "Sample document on 2 columns")]
    public void ConvertWordToMarkdown_WhenOK(string documentPath)
    {
        // Arrange
        var parser = new PdfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, documentPath);

        // Act
        var result = parser.ToMarkdown(filePath, new ExtractOptions()
        {
            ExtractTables = true,
            ExtractImages = true
        });

        // Assert
        Assert.IsNotNull(result.Output);
    }
}