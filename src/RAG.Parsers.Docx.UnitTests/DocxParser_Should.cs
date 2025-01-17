using Microsoft.Extensions.Logging;

namespace RAG.Parsers.Docx.UnitTests;

[TestClass]
public class DocxParser_Should
{
    private const string TestFilesDirectory = "TestFiles";
    private const string TestFileName = "Sample.docx";

    [TestMethod]
    public void ConvertWordToMarkdown_WhenOK() 
    {
        // Arrange
        using var loggerFactory = LoggerFactory.Create(loggingBuilder => loggingBuilder
                    .SetMinimumLevel(LogLevel.Trace)
                    .AddConsole());
        var parser = new DocxParser(loggerFactory.CreateLogger<DocxParser>());
        var filePath = GetTestFilePath();

        // Act
        var result = parser.DocToMarkdownWithContext(filePath, new Models.ExtractOptions()
        {  
            ExtractTables = true,
            ExtractImages = true
        });

        // Assert
        Assert.IsNotNull(result.Output);
    }

    /// <summary>
    /// Constructs the full file path for the test document.
    /// </summary>
    /// <returns>The absolute path to the test file.</returns>
    private string GetTestFilePath()
    {
        return Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, TestFileName);
    }

    /// <summary>
    /// Counts occurrences of a specific substring in the given text.
    /// </summary>
    /// <param name="text">The text to search.</param>
    /// <param name="substring">The substring to count.</param>
    /// <returns>The count of occurrences.</returns>
    private int CountOccurrences(string text, string substring)
    {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(substring))
            return 0;

        return text.Split(new[] { substring }, StringSplitOptions.None).Length - 1;
    }

    [TestMethod]
    public void ConvertWordToMarkdownWithContext_WhenOK()
    {
        // Arrange
        using var loggerFactory = LoggerFactory.Create(loggingBuilder => loggingBuilder
                    .SetMinimumLevel(LogLevel.Trace)
                    .AddConsole());
        var parser = new DocxParser(loggerFactory.CreateLogger<DocxParser>());
        var filePath = GetTestFilePath();
        const int expectedImageCount = 2;

        // Act
        var result = parser.DocToMarkdownWithContext(filePath, new Models.ExtractOptions()
        {
            ExtractTables = true,
            ExtractImages = true
        });

        // Assert
        Assert.IsNotNull(result?.Output, "The output from DocToMarkdownWithContext should not be null.");
        Assert.AreEqual(expectedImageCount, result.Images.Count, $"The output should contain {expectedImageCount} distinct images.");
    }

    [TestMethod]
    public void CheckImageFormat_WhenJpegExists()
    {
        // Arrange
        using var loggerFactory = LoggerFactory.Create(loggingBuilder => loggingBuilder
                    .SetMinimumLevel(LogLevel.Trace)
                    .AddConsole());
        var parser = new DocxParser(loggerFactory.CreateLogger<DocxParser>());
        var filePath = GetTestFilePath();
        const string expectedImageFormat = "jpeg";
        const int expectedJpegCount = 1;

        // Act
        var result = parser.DocToMarkdownWithContext(filePath, new Models.ExtractOptions()
        {
            ExtractTables = true,
            ExtractImages = true
        });

        // Assert
        Assert.IsNotNull(result?.Output, "The output from DocToMarkdownWithContext should not be null.");

        var jpegCount = result.Images.Count(image => image.Format.Equals(expectedImageFormat, StringComparison.OrdinalIgnoreCase));
        Assert.AreEqual(expectedJpegCount, jpegCount, $"The output should contain {expectedJpegCount} images in {expectedImageFormat} format.");
    }

    [TestMethod]
    public void CheckImageTagCount()
    {
        // Arrange
        using var loggerFactory = LoggerFactory.Create(loggingBuilder => loggingBuilder
                    .SetMinimumLevel(LogLevel.Trace)
                    .AddConsole());
        var parser = new DocxParser(loggerFactory.CreateLogger<DocxParser>());
        var filePath = GetTestFilePath();
        const int expectedImageCount = 2;

        // Act
        var result = parser.DocToMarkdownWithContext(filePath, new Models.ExtractOptions()
        {
            ExtractTables = true,
            ExtractImages = true
        });

        // Assert
        Assert.IsNotNull(result?.Output, "The output from DocToMarkdownWithContext should not be null.");

        var imageCount = CountOccurrences(result.Output, "![image]");
        Assert.AreEqual(expectedImageCount, imageCount, $"The output should contain {expectedImageCount} image tags.");
    }
}