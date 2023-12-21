using System;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace My.Converters.ToMarkdown
{
    /// <summary>
    /// Excel Decoder to Markdown
    /// </summary>
    public class MsExcelDecoder(bool withQuotes = true,
                          string? worksheetNumberTemplate = null)
    {
        #region Properties

        private const string DefaultSheetNumberTemplate = "\n# Worksheet \"{name}\"\n";
        private const string DefaultCellBalise = "|";

        private readonly bool _withQuotes = withQuotes;
        private readonly string _worksheetNumberTemplate = worksheetNumberTemplate ?? DefaultSheetNumberTemplate;

        #endregion

        #region Public Methods

        /// <summary>
        /// Read file and open it
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string ExcelToText(string filePath)
        {
            // Open file
            using var stream = File.OpenRead(filePath);

            // Convert file
            return ExcelToText(stream);
        }

        /// <summary>
        /// Convert excel document stream to text
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public string ExcelToText(Stream data)
        {
            using var workbook = new XLWorkbook(data);
            var sb = new StringBuilder();

            foreach (var worksheet in workbook.Worksheets)
            {
                sb.AppendLine(_worksheetNumberTemplate.Replace("{name}", $"{worksheet.Name}"));

                //var columnCount = worksheet.RangeUsed().ColumnCount();                

                var columnCount = 0;
                var firstRow = true;
                
                foreach (IXLRangeRow? row in worksheet.RangeUsed().RowsUsed())
                {
                    //if (row is null) { continue; }                    

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

        const int ColumnBase = 26;
        const int DigitMax = 7; // cell(log26(Int32.Max))
        const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public string IndexToColumn(int index)
        {
            if (index <= 0)
                throw new IndexOutOfRangeException("index must be a positive number");

            if (index <= ColumnBase)
                return Digits[index - 1].ToString();

            var sb = new StringBuilder().Append(' ', DigitMax);
            var current = index;
            var offset = DigitMax;
            while (current > 0)
            {
                sb[--offset] = Digits[--current % ColumnBase];
                current /= ColumnBase;
            }
            return sb.ToString(offset, DigitMax - offset);
        }

        #endregion
    }
}
