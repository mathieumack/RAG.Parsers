using System;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace RAG.Parsers.Xlsx
{
    /// <summary>
    /// Excel Decoder to Markdown
    /// </summary>
    public class XlsxParser(bool withQuotes = true,
                                string? worksheetNumberTemplate = null)
    {
        #region Properties

        private const string DefaultSheetNumberTemplate = "\n# Worksheet \"{name}\"\n";
        private const char DefaultCellBalise = '|';

        private readonly bool _withQuotes = withQuotes;
        private readonly string _worksheetNumberTemplate = worksheetNumberTemplate ?? DefaultSheetNumberTemplate;

        #endregion

        #region Public Methods

        /// <summary>
        /// Read file and open it
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string ExcelToMarkdown(string filePath)
        {
            // Open file
            using var stream = File.OpenRead(filePath);

            // Convert file
            return ExcelToMarkdown(stream);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Convert excel document stream to text
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        private string ExcelToMarkdown(Stream data)
        {
            using var workbook = new XLWorkbook(data);
            var sb = new StringBuilder();

            foreach (var worksheet in workbook.Worksheets)
            {
                sb.AppendLine(_worksheetNumberTemplate.Replace("{name}", $"{worksheet.Name}"));


                var columnCount = 0;
                var firstRow = true;
                
                foreach (IXLRangeRow? row in worksheet.RangeUsed().RowsUsed())
                {

                    if (firstRow)
                    {
                        var rowString = "|";
                        foreach (var cell in row.Cells())
                        {
                            rowString += "|";
                            rowString += cell.Address.ColumnLetter;
                            columnCount++;
                        }

                        rowString += "|";
                        sb.AppendLine(rowString);

                        rowString = "|";
                        var headerRowSeparator = "---|";
                        rowString += new StringBuilder(headerRowSeparator.Length * (columnCount + 1))
                                        .Insert(0, headerRowSeparator, (columnCount + 1))
                                        .ToString();

                        sb.AppendLine(rowString);
                        firstRow = false;
                    }

                    var cells = row.CellsUsed().ToList();

                    sb.Append(DefaultCellBalise);

                    var firstColumn = true;
                    foreach (var cell in row.Cells())
                    {
                        if (firstColumn)
                        {
                            sb.Append("**");
                            sb.Append(cell.Address.RowNumber);
                            sb.Append("**|");
                            firstColumn = false;
                        }

                        if (_withQuotes && cell is { Value.IsText: true })
                            // TODO : Manage strange char
                            sb.Append(cell.Value.GetText().Replace("\"", "\"\""));
                        else
                            sb.Append(cell.Value);
                        
                        sb.Append(DefaultCellBalise);
                    }
                    
                    sb.Append(DefaultCellBalise);

                    sb.AppendLine();
                }
            }

            return sb.ToString().Trim();
        }

        #endregion
    }
}
