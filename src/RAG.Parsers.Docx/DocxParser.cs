using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using RAG.Parsers.Docx.Models;
using RAG.Parsers.Docx.Models.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RAG.Parsers.Docx;

/// <summary>
/// Docx Decoder to Markdown
/// </summary>
public class DocxParser : IDisposable
{
    private readonly ILogger<DocxParser> logger;
    private const int DefaultMemoryStreamCapacity = 8192;

    public DocxParser(ILogger<DocxParser> logger)
    {
        this.logger = logger;
    }

    #region Public Methods

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public ExtractOutput DocToMarkdownWithContext(string filePath, ExtractOptions options)
    {
        // Open file
        using var stream = File.OpenRead(filePath);

        // Convert file
        return DocToMarkdownWithContext(stream, options);
    }

    #endregion

    #region Private Methods

    #region Explore document

    /// <summary>
    /// Convert word document stream to text
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public ExtractOutput DocToMarkdownWithContext(Stream data, ExtractOptions options)
    {
        // Get file from stream
        var context = new ExtractOutput()
        {
            Hyperlinks = new(),
            DictionaryStyles = new(),
            Images = new()
        };

        // Handle null or Stream.Null
        if (data == null || data == Stream.Null)
        {
            logger.LogWarning("Empty file provided.");
            context.Output = string.Empty;
            return context;
        }

        // For non-seekable streams, wrap in MemoryStream to avoid consuming bytes during empty check
        Stream streamToUse = data;
        MemoryStream memoryStream = null;
        if (!data.CanSeek)
        {
            // Use initial capacity to avoid buffer reallocations during copy
            memoryStream = new MemoryStream(capacity: DefaultMemoryStreamCapacity);
            data.CopyTo(memoryStream);
            memoryStream.Position = 0;
            streamToUse = memoryStream;
        }

        try
        {
            // Check if stream is empty
            if (streamToUse.Length == 0)
            {
                logger.LogWarning("Empty file provided.");
                context.Output = string.Empty;
                return context;
            }

            using var wordprocessingDocument = WordprocessingDocument.Open(streamToUse, false);

            StringBuilder sb = new();
            MainDocumentPart? mainPart = wordprocessingDocument.MainDocumentPart ??
                throw new InvalidOperationException("The main document part is missing.");
            Body? body = mainPart.Document.Body ??
                throw new InvalidOperationException("The document body is missing.");

            // Populate hyperlinks and styles
            context.Hyperlinks = GetAllHyperlinks(mainPart);
            context.DictionaryStyles = GetAllStyles(mainPart);

            // Récupérer les commentaires du document
            var commentsPart = mainPart.GetPartsOfType<WordprocessingCommentsPart>().FirstOrDefault();
            var allComments = commentsPart?.Comments?.Elements<Comment>().ToList() ?? new List<Comment>();
            var commentMap = new Dictionary<string, CommentInfo>();
            int commentCounter = 1;

            // Explore file
            var parts = mainPart.Document.Descendants().FirstOrDefault();
            if (parts != null)
                // Explore all elements in file
                foreach (var node in parts.ChildElements)
                {
                    if (node is Paragraph paragraph)
                        // Process Text and paragraph
                        ProcessParagraph(mainPart, paragraph, context, options, ref sb, allComments, commentMap, ref commentCounter);
                    else if (node is Table table && options.ExtractTables)
                        // Process Table
                        ProcessTable(mainPart, table, context, ref sb);
                }

            // Add comment section at the end of the document
            if (commentMap.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("> Comments");
                foreach (var kvp in commentMap.OrderBy(x => x.Value.Index))
                {
                    sb.AppendLine($"> ({kvp.Value.Index}) {kvp.Value.Text}");
                }
            }

            // Return text generated
            var textContent = sb.ToString().Trim();
            context.Output = textContent;

            return context;
        }
        finally
        {
            // Clean up memory stream if we created one
            memoryStream?.Dispose();
        }
    }

    #endregion

    #region Paragraph

    /// <summary>
    /// Process paragraph
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="context"></param>
    /// <param name="sb"></param>
    private void ProcessParagraph(MainDocumentPart mainPart,
                                  Paragraph paragraph,
                                  ExtractOutput context,
                                  ExtractOptions options,
                                  ref StringBuilder sb,
                                  List<Comment> allComments,
                                  Dictionary<string, CommentInfo> commentMap,
                                  ref int commentCounter)
    {
        var stringToAdd = "";

        // Detect and manage TOC
        if (IsParagraphLinkedToTableOfContent(paragraph, context))
            return;

        // Detect and manage title
        if (IsPaagraphLinkedToTitle(paragraph, context))
        {
            stringToAdd += GetTitle(paragraph, context);
            sb.AppendLine(stringToAdd);
            return;
        }

        // Explore all sub elements

        foreach (var child in paragraph.ChildElements)
        {
            if (child is DeletedRun && options.ExtractRevisionContent)
            {
                stringToAdd += GetDeletedText((DeletedRun)child);
                continue;
            }

            // Empty elements, bypass
            if (string.IsNullOrEmpty(child.InnerText))
                continue;

            // Extract text honoring revisions (deleted) and hyperlinks
            stringToAdd += ExtractTextWithRevisions(child, options, context.Hyperlinks);
        }

        // Add comment index and details if any
        var commentIndices = new List<int>();
        var commentInfos = new List<CommentInfo>();
        foreach (var commentRef in paragraph.Descendants<CommentRangeStart>())
        {
            var commentId = commentRef.Id.Value;
            if (!commentMap.ContainsKey(commentId))
            {
                var comment = allComments.FirstOrDefault(c => c.Id.Value == commentId);
                if (comment != null)
                {
                    var author = comment.Author ?? "";
                    var date = comment.Date != null ? comment.Date.Value.ToString("yyyy-MM-dd") : "";
                    var info = new CommentInfo { Index = commentCounter, Text = comment.InnerText, Author = author, Date = date };
                    commentMap[commentId] = info;
                    commentIndices.Add(commentCounter);
                    commentInfos.Add(info);
                    commentCounter++;
                }
            }
            else
            {
                var info = commentMap[commentId];
                commentIndices.Add(info.Index);
                commentInfos.Add(info);
            }
        }
        if (commentIndices.Count > 0)
        {
            stringToAdd += string.Join("", commentIndices.Select(i => $"({i})"));
        }

        if (!string.IsNullOrEmpty(stringToAdd))
        {
            sb.AppendLine(stringToAdd);
            sb.AppendLine();
        }

        // Affichage des commentaires juste sous le texte lié
        foreach (var info in commentInfos)
        {
            sb.AppendLine($"> ({info.Index}) : {info.Author} ({info.Date}) : {info.Text}");
            sb.AppendLine();
        }

        if (commentInfos.Any())
            sb.AppendLine();

        // Now add drawing elements on ths paragraph:
        if (options.ExtractImages)
        {
            foreach (var drawing in paragraph.Descendants<Drawing>())
            {
                ProcessDrawing(drawing, mainPart, context, ref sb);
                sb.AppendLine();
            }
        }
    }

    /// <summary>
    /// Processes a Drawing element, extracts the image, and appends the image reference to the md output.
    /// </summary>
    /// <param name="drawing">The Drawing element to process.</param>
    /// <param name="mainPart">The main document part.</param>
    /// <param name="context">The context containing extracted information.</param>
    /// <param name="sb">The StringBuilder to append the image reference to.</param>
    private void ProcessDrawing(Drawing drawing, MainDocumentPart mainPart, ExtractOutput context, ref StringBuilder sb)
    {
        if (!TryGetImagePart(drawing, mainPart, out var imagePart, out var imageUri, out var imageFormat))
        {
            logger.LogInformation("DocxParser-No image part found for the given drawing.");

            return;
        }

        try
        {
            var imageBytes = GetImageBytes(imagePart);

            var id = $"{Guid.NewGuid()}.{imageFormat}";

            var raw = $"![image](data:image/{imageFormat};{id})";
            sb.Append(raw);

            context.Images.Add(new Models.ImageRef
            {
                Id = id,
                Format = imageFormat,
                MarkdownRaw = raw,
                RawBytes = imageBytes
            });
        }
        catch (Exception ex)
        {
            logger.LogInformation($"DocxParser-Error processing image with URI `{imageUri}`: {ex.Message}");
        }

    }

    /// <summary>
    /// Processes a table element, converting it to Markdown format and appending it to the output.
    /// </summary>
    /// <param name="mainPart">The main document part.</param>
    /// <param name="table">The table element to process.</param>
    /// <param name="context">The context containing extracted information.</param>
    /// <param name="sb">The StringBuilder to append the Markdown output to.</param>
    private void ProcessTable(MainDocumentPart mainPart, Table table, ExtractOutput context, ref StringBuilder sb)
    {
        var cellMatrix = BuildCellMatrix(table);
        var columnCount = cellMatrix.Max(row => row.Count);
        var isFirstRow = true;

        foreach (var matrixRow in cellMatrix)
        {
            var rowToBuild = new StringBuilder();

            // Add blank header if first row isn't header
            if (isFirstRow && matrixRow.Any() && matrixRow[0].Cell != null)
            {
                AppendBlankHeader(sb, columnCount);
                isFirstRow = false;
            }

            for (int colIndex = 0; colIndex < columnCount; colIndex++)
            {
                var cellInfo = matrixRow.Count > colIndex ? matrixRow[colIndex] : null;

                if (cellInfo?.IsOccupied ?? false)
                {
                    rowToBuild.Append('|');
                    if (cellInfo.Cell != null)
                    {
                        ProcessCellContent(cellInfo.Cell, mainPart, context, ref rowToBuild);

                        if (cellInfo.VerticalMerge == MergeType.Continue)
                        {
                            rowToBuild.Append("^^");
                        }
                    }
                }
                else if (cellInfo?.HorizontalMerge == MergeType.Continue)
                {
                    rowToBuild.Append("|<<");
                }
                else
                {
                    rowToBuild.Append('|');
                }
            }

            sb.AppendLine(rowToBuild.ToString() + "|");

            if (isFirstRow)
            {
                sb.AppendLine(BuildTableHeaderSeparator(columnCount));
                isFirstRow = false;
            }
        }
        sb.AppendLine();
    }

    /// <summary>
    /// Appends a blank header row to the output for a table with the specified number of columns.
    /// </summary>
    /// <param name="sb">The StringBuilder to append the blank header to.</param>
    /// <param name="columnCount">The number of columns in the table.</param>
    private void AppendBlankHeader(StringBuilder sb, int columnCount)
    {
        var headerRow = "|";
        for (int i = 0; i < columnCount; i++)
        {
            headerRow += " |";
        }
        sb.AppendLine(headerRow);
        sb.AppendLine(BuildTableHeaderSeparator(columnCount));
    }

    /// <summary>
    /// Processes the content of a table cell, converting it to Markdown format and appending it to the output.
    /// </summary>
    /// <param name="cell">The table cell to process.</param>
    /// <param name="mainPart">The main document part.</param>
    /// <param name="context">The context containing extracted information.</param>
    /// <param name="sb">The StringBuilder to append the Markdown output to.</param>
    private void ProcessCellContent(TableCell cell, MainDocumentPart mainPart, ExtractOutput context, ref StringBuilder sb)
    {
        // Build cell content while respecting revisions (deleted text) and hyperlinks
        var cellContent = new StringBuilder();
        foreach (var child in cell.ChildElements)
        {
            if (string.IsNullOrEmpty(child.InnerText))
                continue;

            cellContent.Append(ExtractTextWithRevisions(child, new ExtractOptions(), context.Hyperlinks));
        }

        sb.Append(cellContent.ToString());
        foreach (var drawing in cell.Descendants<Drawing>())
        {
            ProcessDrawing(drawing, mainPart, context, ref sb);
        }
    }

    /// <summary>
    /// Builds a matrix of CellInfo objects representing the structure of the table, including merged cells.
    /// </summary>
    /// <param name="table">The table element to analyze.</param>
    /// <returns>A list of lists of CellInfo objects representing the table structure.</returns>
    private List<List<CellInfo>> BuildCellMatrix(Table table)
    {
        var matrix = new List<List<CellInfo>>();
        int rowIndex = 0;

        foreach (var row in table.Elements<TableRow>())
        {
            var matrixRow = new List<CellInfo>();
            int colIndex = 0;

            while (colIndex < matrix.FirstOrDefault()?.Count || row.Elements<TableCell>().Any())
            {
                while (colIndex < matrixRow.Count && matrixRow[colIndex].IsOccupied)
                {
                    colIndex++;
                }

                if (!row.Elements<TableCell>().Any()) break;

                var cell = row.Elements<TableCell>().First();
                var cellInfo = AnalyzeCell(cell, rowIndex, colIndex);

                for (int i = 0; i < cellInfo.GridSpan; i++)
                {
                    matrixRow.Add(new CellInfo
                    {
                        Cell = i == 0 ? cell : null,
                        HorizontalMerge = i == 0 ? MergeType.First : MergeType.Continue,
                        IsOccupied = true
                    });
                }

                if (cellInfo.VerticalMerge == MergeType.First)
                {
                    for (int i = 1; i < cellInfo.VerticalSpan; i++)
                    {
                        while (matrix.Count <= rowIndex + i)
                        {
                            matrix.Add(new List<CellInfo>());
                        }
                        var targetRow = matrix[rowIndex + i];
                        while (targetRow.Count <= colIndex)
                        {
                            targetRow.Add(new CellInfo());
                        }
                        targetRow[colIndex] = new CellInfo
                        {
                            VerticalMerge = MergeType.Continue,
                            IsOccupied = true
                        };
                    }
                }

                row.RemoveChild(cell);
                colIndex += cellInfo.GridSpan;
            }

            matrix.Add(matrixRow);
            rowIndex++;
        }

        return matrix;
    }

    /// <summary>
    /// Analyzes a table cell to determine its merge properties and span.
    /// </summary>
    /// <param name="cell">The table cell to analyze.</param>
    /// <param name="rowIndex">The row index of the cell in the table.</param>
    /// <param name="colIndex">The column index of the cell in the table.</param>
    /// <returns>A CellInfo object containing the cell's properties.</returns>
    private CellInfo AnalyzeCell(TableCell cell, int rowIndex, int colIndex)
    {
        var cellInfo = new CellInfo { Cell = cell };

        var tcPr = cell.GetFirstChild<TableCellProperties>();
        if (tcPr != null)
        {
            var hMerge = tcPr.GetFirstChild<HorizontalMerge>();
            if (hMerge != null)
            {
                cellInfo.HorizontalMerge = hMerge.Val == null || hMerge.Val == "continue"
                    ? MergeType.Continue
                    : MergeType.First;
            }

            var vMerge = tcPr.GetFirstChild<VerticalMerge>();
            if (vMerge != null)
            {
                cellInfo.VerticalMerge = vMerge.Val == null || vMerge.Val == "continue"
                    ? MergeType.Continue
                    : MergeType.First;
            }

            var gridSpan = tcPr.GetFirstChild<GridSpan>();
            if (gridSpan != null)
            {
                cellInfo.GridSpan = gridSpan.Val.Value;
            }
        }

        return cellInfo;
    }

    /// <summary>
    /// Build TableHeader Separator
    /// </summary>
    /// <param name="numberOfColumn"></param>
    /// <returns></returns>
    private static string BuildTableHeaderSeparator(int numberOfColumn)
    {
        var row = "|";
        var headerRowSeparator = "---|";
        row += new StringBuilder(headerRowSeparator.Length * numberOfColumn)
                        .Insert(0, headerRowSeparator, numberOfColumn)
                        .ToString();

        return row;
    }

    #endregion

    #region Hyperlinks and Styles

    /// <summary>
    /// Get all hyperlinks in document
    /// </summary>
    /// <param name="mainDocument"></param>
    private List<HyperlinkRelationship> GetAllHyperlinks(MainDocumentPart mainDocument)
    {
        return mainDocument.HyperlinkRelationships.Where(x => x.IsExternal).ToList();
    }

    /// <summary>
    /// Get Hyperlinkg
    /// </summary>
    /// <param name="element"></param>
    /// <param name="hyperlinks"></param>
    /// <returns></returns>
    private static string GetHyperlink(Hyperlink element, List<HyperlinkRelationship> hyperlinks)
    {
        var stringToReturn = "";

        // Generate text for hyperlink with decoration
        try
        {
            stringToReturn += "[";
            stringToReturn += element.InnerText;
            stringToReturn += "]";

            var hyperlink = hyperlinks?.FirstOrDefault(x => x.Id == element.Id);

            // Only link with external (URL) hyperlink
            if (hyperlink is not null && hyperlink.IsExternal)
            {
                stringToReturn += "(";
                stringToReturn += hyperlink.Uri.AbsoluteUri;
                stringToReturn += ")";
            }
        }
        catch (Exception)
        {
            return stringToReturn;
        }

        return stringToReturn;
    }

    /// <summary>
    /// Get all styles in document
    /// </summary>
    /// <param name="mainDocument"></param>
    private Dictionary<string, (bool isHeadingStyle, bool isTOCStyle)> GetAllStyles(MainDocumentPart mainDocument)
    {
        var styles = mainDocument.StyleDefinitionsPart.Styles.StylesPart.Styles;

        var dictionaryStyles = new Dictionary<string, (bool isHeadingStyle, bool isTOCStyle)>();

        foreach (Style style in styles.ChildElements.Where(x => x.GetType() == typeof(Style)).Cast<Style>())
        {
            var hasStyleName = style.StyleName != null;
            bool isHeading = false;
            bool isTOCStyle = false;
            if (hasStyleName)
            {
                isHeading = style.StyleName.Val.Value.Contains("heading");
                isTOCStyle = style.StyleName.Val.Value.Contains("toc");
            }
            dictionaryStyles.Add(style.StyleId, (isHeading, isTOCStyle));
        }

        return dictionaryStyles;
    }

    #endregion

    #region Labels and Decoration

    private static string GetDeletedText(DeletedRun element)
    {
        var stringToReturn = "";
        try
        {
            stringToReturn += $"~~(revision : {element.Author} - {element.Date} : {element.InnerText.Trim()})~~";
        }
        catch (Exception)
        {
            return stringToReturn;
        }
        return stringToReturn;
    }

    /// <summary>
    /// Get text and styling associated
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static string GetLabelAndDecoration(OpenXmlElement element, ExtractOptions options)
    {
        // Dispatcher
        return element.FirstChild switch
        {
            RunProperties => GetLabelAndDecorationRunChild(element),
            Text => GetLabelAndDecorationTextChild(element),
            _ => GetLabelAndDecorationTextChild(element),
        };
    }

    /// <summary>
    /// Get text and styling for a RunProperties element
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static string GetLabelAndDecorationRunChild(OpenXmlElement element)
    {
        var stringToReturn = "";

        try
        {
            var runProperties = (RunProperties)element.FirstChild;

            // Detect decoration
            var balise = "";
            if (runProperties.Bold is not null)
                balise += "**";

            if (runProperties.Italic is not null)
                balise += "*";

            var textToAdd = element.InnerText;

            // Detect details - whitespace before/after
            var whitespacePreWord = false;
            if (textToAdd.StartsWith(" "))
                whitespacePreWord = true;

            var whitespacePostWord = false;
            if (textToAdd.EndsWith(" "))
                whitespacePostWord = true;

            // Clean
            textToAdd = textToAdd.TrimStart(' ');
            textToAdd = textToAdd.TrimEnd(' ');

            // Adapt return with details
            if (whitespacePreWord)
                stringToReturn += " ";
            stringToReturn += balise;
            stringToReturn += textToAdd;
            stringToReturn += balise;
            if (whitespacePostWord)
                stringToReturn += " ";
        }
        catch (Exception)
        {
            return stringToReturn;
        }

        return stringToReturn;
    }

    /// <summary>
    /// Get text and styling for a Text element
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static string GetLabelAndDecorationTextChild(OpenXmlElement element)
    {
        var stringToReturn = "";

        try
        {
            var runProperties = (Text)element.FirstChild;

            stringToReturn = element.InnerText;
        }
        catch (Exception)
        {
            return stringToReturn;
        }

        return stringToReturn;
    }

    /// <summary>
    /// Extracts text from the element while preserving deleted revision text and hyperlinks when requested.
    /// </summary>
    /// <param name="element">Element to extract.</param>
    /// <param name="options">Extraction options (controls revision content extraction).</param>
    /// <param name="hyperlinks">List of hyperlinks for resolving links.</param>
    /// <returns>Extracted string.</returns>
    private string ExtractTextWithRevisions(OpenXmlElement element, ExtractOptions options, List<HyperlinkRelationship> hyperlinks)
    {
        var sb = new StringBuilder();
        var children = element.ChildElements.ToList();
        foreach (var child in children)
        {
            if (child is Hyperlink hyperlinkElem)
            {
                sb.Append(GetHyperlink(hyperlinkElem, hyperlinks));
            }
            else if (options != null && options.ExtractRevisionContent && child is DeletedRun deleted)
            {
                sb.Append(GetDeletedText(deleted));
            }
            else
            {
                sb.Append(GetLabelAndDecoration(child, options));
            }
        }

        return sb.ToString();
    }

    #endregion

    #region TOC

    /// <summary>
    /// Detect if element is from Table of Content
    /// </summary>
    /// <param name="paragraph"></param>
    /// <returns></returns>
    private bool IsParagraphLinkedToTableOfContent(Paragraph paragraph, ExtractOutput context)
    {
        try
        {
            var paragraphProperties = (ParagraphProperties)paragraph.FirstChild;
            var paragraphStyleId = (ParagraphStyleId)paragraphProperties.FirstChild;

            if (context.DictionaryStyles.ContainsKey(paragraphStyleId.Val.Value))
                return context.DictionaryStyles[paragraphStyleId.Val.Value].isTOCStyle;
        }
        catch (Exception)
        {
        }

        return false;
    }

    #endregion

    #region Titles

    /// <summary>
    /// Detect if element is a title
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="context"></param>
    /// <returns></returns>
    private bool IsPaagraphLinkedToTitle(Paragraph paragraph, ExtractOutput context)
    {
        try
        {
            var paragraphProperties = (ParagraphProperties)paragraph.FirstChild;
            var paragraphStyleId = (ParagraphStyleId)paragraphProperties.FirstChild;

            if (context.DictionaryStyles.ContainsKey(paragraphStyleId.Val.Value))
                return context.DictionaryStyles[paragraphStyleId.Val.Value].isHeadingStyle;
        }
        catch (Exception)
        {
        }

        return false;
    }

    /// <summary>
    /// Get title level and associated decoration
    /// </summary>
    /// <param name="paragraph"></param>
    /// <returns></returns>
    private string GetTitle(Paragraph paragraph, ExtractOutput context)
    {
        var stringToReturn = "";
        var paragraphProperties = (ParagraphProperties)paragraph.FirstChild;
        var paragraphStyleId = (ParagraphStyleId)paragraphProperties.FirstChild;

        // Get style from dictionary
        if (IsPaagraphLinkedToTitle(paragraph, context))
        {
            // Get level and adapt for markdown
            string lastChar = paragraphStyleId.Val.Value[^1..];
            int.TryParse(lastChar, out var titleLvl);
            titleLvl++;

            stringToReturn += new string('#', titleLvl);
            stringToReturn += " ";
            stringToReturn += paragraph.InnerText;
        }

        return stringToReturn;
    }

    #endregion

    #region Image

    /// <summary>
    /// Get image content
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static string GetImageContent(OpenXmlElement element)
    {
        // Dispatcher
        return element.FirstChild switch
        {
            RunProperties => GetLabelAndDecorationRunChild(element),
            Text => GetLabelAndDecorationTextChild(element),
            _ => GetLabelAndDecorationTextChild(element),
        };
    }

    /// <summary>
    /// Gets the image format from the ImagePart.
    /// </summary>
    /// <param name="imagePart">The image part.</param>
    /// <returns>The image format as a string.</returns>
    private string GetImageFormat(ImagePart imagePart)
    {
        var contentTypeParts = imagePart.ContentType.Split('/');

        return contentTypeParts.Length > 1 ? contentTypeParts[1] : null;
    }

    /// <summary>
    /// Tries to get the ImagePart, image URI, and image format from a Drawing element.
    /// </summary>
    /// <param name="drawing">The drawing element.</param>
    /// <param name="mainPart">The main document part.</param>
    /// <param name="imagePart">The image part.</param>
    /// <param name="imageUri">The image URI.</param>
    /// <param name="imageFormat">The image format.</param>
    /// <returns>True if the image part is found; otherwise, false.</returns>
    bool TryGetImagePart(Drawing drawing, MainDocumentPart mainPart, out ImagePart imagePart, out string imageUri, out string imageFormat)
    {
        imagePart = null;
        imageUri = null;
        imageFormat = null;

        var graphic = drawing.Inline?.Graphic ?? drawing.Anchor.Descendants<DocumentFormat.OpenXml.Drawing.Graphic>().FirstOrDefault();
        var image = graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();

        if (image == null)
            return false;

        imageUri = image.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(imageUri))
            return false;

        imagePart = mainPart.GetPartById(imageUri) as ImagePart;
        if (imagePart == null)
            return false;

        imageFormat = GetImageFormat(imagePart);

        return !string.IsNullOrEmpty(imageFormat);
    }

    /// <summary>
    /// Gets the image bytes from an ImagePart.
    /// </summary>
    /// <param name="imagePart">The image part.</param>
    /// <returns>The image bytes as a byte array.</returns>
    byte[] GetImageBytes(ImagePart imagePart)
    {
        using (var imageStream = imagePart.GetStream())
        using (var memoryStream = new MemoryStream())
        {
            imageStream.CopyTo(memoryStream);
            return memoryStream.ToArray();
        }
    }
    #endregion

    #endregion

    public void Dispose()
    {
        // Nothing to release
        GC.SuppressFinalize(this);
    }

    private class CommentInfo
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public string Author { get; set; }
        public string Date { get; set; }
    }
}