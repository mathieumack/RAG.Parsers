namespace My.Converters.ToMarkdown.UnitTests
{
    [TestClass]
    public class MsExcelDecoder_Should
    {
        [TestMethod]
        public void TestMethod1()
        {
            // Arrange
            var excelParser = new MsExcelDecoder();
            var templatePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/SimpleSheet.xlsx");

            // Act
            var result = excelParser.ExcelToText(templatePath);

            // Assert
            Assert.IsNotNull(result);   
        }

        [TestMethod]
        public void TestMethod2()
        {
            // Arrange
            var excelParser = new MsExcelDecoder();

            // Act
            var result = excelParser.IndexToColumn(50);

            // Assert
            Assert.IsNotNull(result);
        }
    }
}