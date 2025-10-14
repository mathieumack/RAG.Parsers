namespace RAG.Parsers.VTT.UnitTests;

[TestClass]
public class VttParserTests
{
    [TestMethod]
    [DataRow("TestFiles/test large vtt.vtt", DisplayName = "Sample document")]
    public void ToMarkdown_ParsesLargeVttFile_CreatesMarkdown(string documentPath)
    {
        // Arrange
        var parser = new VttParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, documentPath);
        using var stream = File.OpenRead(filePath);

        // Act
        var result = parser.ToMarkdown(stream);

        // Assert
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output));
    }
}
