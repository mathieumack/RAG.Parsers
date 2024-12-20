using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using RAG.Parsers.Docx.Models;

namespace RAG.Parsers.Docx;

/// <summary>
/// Docx Decoder to Markdown
/// </summary>
public class DocxParser
{
    #region Public Methods

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public ExtractOutput DocToMarkdownWithContext(string filePath)
    {
        // Open file
        using var stream = File.OpenRead(filePath);

        // Convert file
        return DocToMarkdownWithContext(stream);
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
    public ExtractOutput DocToMarkdownWithContext(Stream data, HashSet<string> supportedFormats = null)
    {
        // Get file from stream
        var context = new ExtractOutput()
        {
            Hyperlinks = new(),
            DictionaryStyles = new(),
            Images = new()
        };

        var wordprocessingDocument = WordprocessingDocument.Open(data, false);
        try
        {
            // Stringbuilder for the output
            StringBuilder sb = new();

            MainDocumentPart? mainPart = wordprocessingDocument.MainDocumentPart ??
                throw new InvalidOperationException("The main document part is missing.");

            Body? body = mainPart.Document.Body ??
                throw new InvalidOperationException("The document body is missing.");

            // Get Hyperlinks and Styles

            // Explore file
            var parts = mainPart.Document.Descendants().FirstOrDefault();
            if (parts != null)
                // Explore all elements in file
                foreach (var node in parts.ChildElements)
                {
                    if (node is Paragraph paragraph)
                        // Process Text and paragraph
                        ProcessParagraph(mainPart, paragraph, context, ref sb);
                    else if (node is Table table)
                        // Process Table
                        ProcessTable(table, ref sb);
                }

            // Return text generated
            var textContent = sb.ToString().Trim();
            context.Output = textContent;

            return context;
        }
        finally
        {
            // Release file
            wordprocessingDocument.Dispose();
        }

        return context;
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
                                  ref StringBuilder sb)
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
            // Empty elements, bypass
            if (string.IsNullOrEmpty(child.InnerText))
                continue;

            // Hyperlink => get link - otherwise => get text and styling
            if (child.GetType() == typeof(Hyperlink))
                stringToAdd += GetHyperlink((Hyperlink)child, context.Hyperlinks);
            else
                stringToAdd += GetLabelAndDecoration(child);
        }

        if (!string.IsNullOrEmpty(stringToAdd))
            sb.AppendLine(stringToAdd);

        // Now add drawing elements on ths paragraph:
        foreach (var drawing in paragraph.Descendants<Drawing>())
        {
            var image = drawing.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
            
            if (image == null) 
                continue;

            var imageUri = image.BlipFill?.Blip?.Embed?.Value;
            
            if (string.IsNullOrEmpty(imageUri) || context.Images.Any(e => e.Id == imageUri))
                continue;

            try
            {
                // Retrieve the ImagePart from the mainPart using the imageUri
                var imagePart = mainPart.GetPartById(imageUri) as ImagePart;
                
                if (imagePart == null)
                    continue;

                // Determine the image format
                var contentTypeParts = imagePart.ContentType.Split('/');
                var imageFormat = contentTypeParts.Length > 1 ? contentTypeParts[1] : null;
                
                if (string.IsNullOrEmpty(imageFormat))
                    continue;

                byte[] imageBytes;
                using (var imageStream = imagePart.GetStream())
                using (var memoryStream = new MemoryStream())
                {
                    imageStream.CopyTo(memoryStream);
                    imageBytes = memoryStream.ToArray();
                }

                context.Images.Add(new Models.ImageRef
                {
                    Id = imageUri,
                    Format = imageFormat,
                    RawBytes = imageBytes
                });

                // Append the image reference to the string builder
                sb.AppendLine($"![image](data:image/{imageFormat};imageRefId,{imageUri})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing image with the following uri `{imageUri}`: {ex.Message}");
            }
        }

    }

    #endregion

    #region Tables

    /// <summary>
    /// Process Table
    /// </summary>
    /// <param name="table"></param>
    /// <param name="sb"></param>
    private static void ProcessTable(Table table, ref StringBuilder sb)
    {
        var firstRow = true;

        // Explore table row in sub elements
        foreach (var row in table.ChildElements.Where(x => x.GetType() == typeof(TableRow)))
        {
            var rowToBuild = "";

            // Get number of column
            var numberOfColumn = row.Where(x => x.GetType() == typeof(TableCell)).Count();

            // Detect if header 
            // If not, add a 'blank' header row
            var childType = row.ChildElements.FirstOrDefault().GetType();
            if (firstRow && childType == typeof(TableCell))
            {
                var headerRow = "|";
                var headerRowSeparator = "-|";
                headerRow += new StringBuilder(headerRowSeparator.Length * numberOfColumn)
                                .Insert(0, headerRowSeparator, numberOfColumn)
                                .ToString();

                sb.AppendLine(headerRow);
                sb.AppendLine(BuildTableHeaderSeparator(numberOfColumn));

                firstRow = false;
            }

            // Explore cells in row
            foreach (var cell in row.Where(x => x.GetType() == typeof(TableCell)))
            {
                rowToBuild += "|";
                rowToBuild += cell.InnerText;
            }
            rowToBuild += "|";

            sb.AppendLine(rowToBuild);

            // Deal with separator needed for markdown
            if (firstRow)
            {
                sb.AppendLine(BuildTableHeaderSeparator(numberOfColumn));

                firstRow = false;
            }
        }

        sb.AppendLine();
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

    /// <summary>
    /// Get text and styling associated
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static string GetLabelAndDecoration(OpenXmlElement element)
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


    #endregion

    #endregion
}