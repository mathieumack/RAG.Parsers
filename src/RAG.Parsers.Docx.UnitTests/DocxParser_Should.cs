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

        // 4 distinct images
        Assert.AreEqual(4, result.Images.Count);
    }

    [TestMethod]
    public void CheckImageFormat()
    {
        // Arrange
        var parser = new DocxParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.docx");

        // Act
        var result = parser.DocToMarkdownWithContext(filePath);

        // Assert
        Assert.IsNotNull(result.Output);

        // Only one jpeg image
        var jpgImages = result.Images.Where(image => image.Format == "jpeg").Count();
        Assert.AreEqual(1, jpgImages);
    }
}