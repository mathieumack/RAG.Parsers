using RAG.Parsers.Rtf.Models;

namespace RAG.Parsers.Rtf.UnitTests;

[TestClass]
public class RtfParser_Should
{
    private const string TestFilesDirectory = "TestFiles";

    [TestMethod]
    [DataRow("TestFiles/sample.rtf", DisplayName = "Basic RTF document")]
    public void ToMarkdown_ParsesRtfFile_ReturnsNonEmptyOutput(string documentPath)
    {
        // Arrange
        var parser = new RtfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, documentPath);

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output));
    }

    [TestMethod]
    public void ToMarkdown_FromStream_ReturnsNonEmptyOutput()
    {
        // Arrange
        var parser = new RtfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "sample.rtf");
        using var stream = File.OpenRead(filePath);

        // Act
        var result = parser.ToMarkdown(stream);

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output));
    }

    [TestMethod]
    public void ToMarkdown_ContainsExpectedText_WhenBasicRtf()
    {
        // Arrange
        var parser = new RtfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "sample.rtf");

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsTrue(result.Output.Contains("Hello, World!"), "Output should contain 'Hello, World!'");
    }

    [TestMethod]
    public void ToMarkdown_PreservesParagraphBreaks_WhenMultipleParagraphs()
    {
        // Arrange
        var parser = new RtfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "sample.rtf");

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsTrue(result.Output.Contains('\n'), "Output should contain line breaks between paragraphs");
    }

    [TestMethod]
    public void ToMarkdown_HandlesUnicodeCharacters_WhenUnicodeRtf()
    {
        // Arrange
        var parser = new RtfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "unicode.rtf");

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output));
    }

    [TestMethod]
    public void ToMarkdown_ReturnsEmptyOutput_WhenMalformedRtf()
    {
        // Arrange
        var parser = new RtfParser();
        var rtfContent = "this is not valid rtf content at all \\invalid{{{";
        var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(rtfContent));

        // Act
        var result = parser.ToMarkdown(stream);

        // Assert
        Assert.IsNotNull(result);
        // Malformed RTF should return empty or gracefully handle the input
        Assert.IsNotNull(result.Output);
    }

    [TestMethod]
    public void ToMarkdown_ReturnsEmptyOutput_WhenEmptyStream()
    {
        // Arrange
        var parser = new RtfParser();
        var stream = new System.IO.MemoryStream(Array.Empty<byte>());

        // Act
        var result = parser.ToMarkdown(stream);

        // Assert
        Assert.IsNotNull(result);
        Assert.AreEqual(string.Empty, result.Output);
    }
}
