using System;
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
    public string ToMarkdown(string filePath, bool withQuotes = true, string? worksheetNumberTemplate = null)
    {
        // Open file
        using var stream = File.OpenRead(filePath);

        // Convert file
        return ToMarkdown(stream);
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
    public string ToMarkdown(Stream data, bool withQuotes = true, string? worksheetNumberTemplate = null)
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

                    if (context.WithQuotes && cell is { Value.IsText: true })
                        // TODO : Manage strange char
                        sb.Append(cell.Value.GetText().Replace("\"", "\"\""));
                    else
                        sb.Append(cell.Value);
                    
                    sb.Append(DocumentContext.DefaultCellBalise);
                }
                
                sb.Append(DocumentContext.DefaultCellBalise);

                sb.AppendLine();
            }
        }

        return sb.ToString().Trim();
    }

    #endregion
}
