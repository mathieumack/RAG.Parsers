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
        Assert.IsTrue(result.Output.Contains("café"), "Output should contain 'café' (escaped \\' sequences decoded)");
        Assert.IsTrue(result.Output.Contains("naïve"), "Output should contain 'naïve' (escaped \\' sequences decoded)");
        Assert.IsTrue(result.Output.Contains("€"), "Output should contain '€' (RTF unicode escape \\u8364 decoded)");
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
        // Malformed RTF should return empty output
        Assert.IsTrue(string.IsNullOrEmpty(result.Output), "Malformed RTF should produce empty output");
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

    [TestMethod]
    public void ToMarkdown_ExtractsImages_WhenRtfContainsEmbeddedImage()
    {
        // Arrange
        var parser = new RtfParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "image.rtf");

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsNotNull(result);
        Assert.AreEqual(1, result.Images.Count, "Output should contain exactly one extracted image");
        var image = result.Images[0];
        Assert.IsFalse(string.IsNullOrEmpty(image.Id), "Image Id should be set");
        Assert.AreEqual("png", image.Format, "Image format should be 'png'");
        Assert.IsTrue(image.RawBytes.Count > 0, "Image should have non-empty raw bytes");
        Assert.IsTrue(result.Output.Contains($"![image](data:image/png;{image.Id})"),
            "Markdown output should contain the image reference");
    }
}
