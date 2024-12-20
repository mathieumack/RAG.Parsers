namespace RAG.Parsers.Docx.UnitTests;

[TestClass]
public class DocxParser_Should
{
    [TestMethod]
    public void ConvertWordToMarkdownWithContext_WhenOK()
    {
        // Arrange
        var parser = new DocxParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.docx");

        // Act
        var result = parser.DocToMarkdownWithContext(filePath);

        // Assert
        Assert.IsNotNull(result.Output);
        // 2 images :
        Assert.AreEqual(2, result.Images.Count);
    }
}