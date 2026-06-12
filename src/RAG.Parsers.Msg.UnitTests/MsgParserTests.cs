using RAG.Parsers.Msg.Models;

namespace RAG.Parsers.Msg.UnitTests;

[TestClass]
public class MsgParserTests
{
    private const string TestFilesDirectory = "TestFiles";

    private MsgParser CreateParser()
    {
        return new MsgParser(Microsoft.Extensions.Logging.Abstractions.NullLogger<MsgParser>.Instance);
    }

    [TestMethod]
    [DataRow("Sample.msg", DisplayName = "Simple message")]
    [DataRow("SampleWithAttachment.msg", DisplayName = "Message with recipient and attachment")]
    public void ParseMsg_FromFilePath_ReturnsNonEmptyOutput(string fileName)
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, fileName);
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions());

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output));
    }

    [TestMethod]
    [DataRow("Sample.msg", DisplayName = "Simple message stream")]
    public void ParseMsg_FromStream_ReturnsNonEmptyOutput(string fileName)
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, fileName);
        using var parser = CreateParser();
        using var stream = File.OpenRead(filePath);

        // Act
        var result = parser.ParseMsg(stream, new ExtractOptions());

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Output));
    }

    [TestMethod]
    public void ParseMsg_ExtractsSubject()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "Sample.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions());

        // Assert
        Assert.AreEqual("Test Email Subject", result.Subject);
    }

    [TestMethod]
    public void ParseMsg_ExtractsSender()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "Sample.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions());

        // Assert
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Headers.From));
        StringAssert.Contains(result.Headers.From, "sender@example.com");
    }

    [TestMethod]
    public void ParseMsg_ExtractsBodyText()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "Sample.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions());

        // Assert
        Assert.IsFalse(string.IsNullOrWhiteSpace(result.Body.Text));
        StringAssert.Contains(result.Body.Text, "This is the test email body.");
    }

    [TestMethod]
    public void ParseMsg_WithRecipients_ExtractsToAddresses()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "SampleWithAttachment.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions { ExtractAttachments = true });

        // Assert
        Assert.IsTrue(result.Headers.To.Count > 0, "Expected at least one To recipient.");
        Assert.IsTrue(result.Headers.To.Any(t => t.Contains("recipient@example.com")));
    }

    [TestMethod]
    public void ParseMsg_WithAttachments_ExtractsAttachmentMetadata()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "SampleWithAttachment.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions { ExtractAttachments = true });

        // Assert
        Assert.AreEqual(1, result.Attachments.Count, "Expected exactly one attachment.");
        var att = result.Attachments[0];
        Assert.AreEqual("attachment.txt", att.FileName);
        Assert.AreEqual("text/plain", att.ContentType);
        Assert.IsTrue(att.Size > 0);
        Assert.IsTrue(att.RawBytes.Count > 0, "RawBytes should be populated.");
    }

    [TestMethod]
    public void ParseMsg_WhenExtractAttachmentsFalse_DoesNotReturnAttachments()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "SampleWithAttachment.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions { ExtractAttachments = false });

        // Assert
        Assert.AreEqual(0, result.Attachments.Count, "Attachments should not be extracted when ExtractAttachments is false.");
    }

    [TestMethod]
    public void ParseMsg_OutputContainsSubjectAndSender()
    {
        // Arrange
        var filePath = Path.Combine(Environment.CurrentDirectory, TestFilesDirectory, "Sample.msg");
        using var parser = CreateParser();

        // Act
        var result = parser.ParseMsg(filePath, new ExtractOptions());

        // Assert
        StringAssert.Contains(result.Output, "Test Email Subject");
        StringAssert.Contains(result.Output, "sender@example.com");
    }
}
