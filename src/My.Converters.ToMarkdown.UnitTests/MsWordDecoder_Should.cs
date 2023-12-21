namespace My.Converters.ToMarkdown.UnitTests
{
    [TestClass]
    public class MsWordDecoder_Should
    {
        [TestMethod]
        public void TestMethod1()
        {
            // Arrange
            var wordParser = new MsWordDecoder();
            var templatePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/sample.dotx");

            // Act
            var result = wordParser.DocToText(templatePath);

            // Assert
            Assert.IsNotNull(result);   
        }
    }
}