namespace RAG.Parsers.Docx.UnitTests
{
    [TestClass]
    public class DocxParser_Should
    {
        [TestMethod]
        public void ConvertWordToMarkdown_WhenOK()
        {
            // Arrange
            var docxParser = new DocxParser();
            var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.docx");

            // Act
            var result = docxParser.DocToMarkdown(filePath);

            // Assert
            Assert.IsNotNull(result);
        }
    }
}