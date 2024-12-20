namespace RAG.Parsers.Pdf.UnitTests;

[TestClass]
public class PdfParser_Should
{
    [TestMethod]
    public void ConvertWordToMarkdown_WhenOK()
    {
        // Arrange
        var parser = new PdfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.pdf");

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsNotNull(result.Output);
    }
}