namespace RAG.Parsers.Xlmx.UnitTests
{
    [TestClass]
    public class XlmxParser_Should
    {
        [TestMethod]
        public void ConvertExcelToMarkdown_WhenOK()
        {
            // Arrange
            var xlmxParser = new XlmxParser();
            var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.xlsx");

            // Act
            var result = xlmxParser.ExcelToMarkdown(filePath);

            // Assert
            Assert.IsNotNull(result);   
        }
    }
}