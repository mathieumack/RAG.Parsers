using System.Globalization;

namespace RAG.Parsers.Xlsx.UnitTests;

[TestClass]
public class XlsxParser_Should
{
    [TestMethod]
    public void ConvertExcelToMarkdown_WhenOK()
    {
        // Arrange
        var parser = new XlsxParser();
        var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.xlsx");

        // Act
        var result = parser.ToMarkdown(filePath);

        // Assert
        Assert.IsNotNull(result);   
    }
}