using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace RAG.Parsers.Xlsx;

/// <summary>
/// Excel Decoder to Markdown
/// </summary>
public class XlsxParser()
{
    #region Public Methods

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="withQuotes"></param>
    /// <param name="worksheetNumberTemplate"></param>
    /// <returns></returns>
    public string ExcelToMarkdown(string filePath, bool withQuotes = true, string? worksheetNumberTemplate = null)
    {
        // Open file
        using var stream = File.OpenRead(filePath);

        // Convert file
        return ExcelToMarkdown(stream, withQuotes, worksheetNumberTemplate);
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Convert excel document stream to text
    /// </summary>
    /// <param name="data"></param>
    /// <param name="withQuotes"></param>
    /// <param name="worksheetNumberTemplate"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public string ExcelToMarkdown(Stream data, bool withQuotes = true, string? worksheetNumberTemplate = null)
    {
        using var workbook = new XLWorkbook(data);
        var sb = new StringBuilder();

        var context = new DocumentContext()
        {
            WithQuotes = withQuotes,
            WorksheetNumberTemplate = worksheetNumberTemplate ?? DocumentContext.DefaultSheetNumberTemplate
        };

        foreach (var worksheet in workbook.Worksheets)
        {
            sb.AppendLine(context.WorksheetNumberTemplate.Replace("{name}", $"{worksheet.Name}"));

            var columnCount = 0;
            var firstRow = true;

            var rangeUsed = worksheet.RangeUsed();
            if (rangeUsed is null) // Worksheet is empty
                continue;

            foreach (IXLRangeRow? row in rangeUsed.RowsUsed())
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

                sb.Append(DocumentContext.DefaultCellBalise);

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

                    if (cell is { Value.IsNumber: true })
                        sb.Append(cell.CachedValue.GetNumber().ToString(CultureInfo.InvariantCulture));
                    else if (cell is { Value.IsDateTime: true })
                        sb.Append(cell.CachedValue.GetDateTime().ToString(CultureInfo.InvariantCulture));
                    else if (cell is { Value.IsBlank: true })
                        sb.Append(cell.CachedValue.GetBlank().ToString());
                    else
                        sb.Append(cell.CachedValue);

                    sb.Append(DocumentContext.DefaultCellBalise);
                }
                
                sb.AppendLine();
            }
        }

        return sb.ToString().Trim();
    }

    #endregion
}
