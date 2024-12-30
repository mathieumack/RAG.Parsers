using System.Text;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using System;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using System.IO;
using RAG.Parsers.Pdf.Models;
using System.Linq;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;

namespace RAG.Parsers.Pdf;

/// <summary>
/// Docx Decoder to Markdown
/// </summary>
public class PdfParser
{
    #region Public Methods

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public ExtractOutput ToMarkdown(string filePath, ExtractOptions options)
    {
        // Open file
        using var stream = File.OpenRead(filePath);

        // Convert file
        return ToMarkdown(stream, options);
    }

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    public ExtractOutput ToMarkdown(Stream data, ExtractOptions options)
    {
        var result = new ExtractOutput()
        {
            Images = new(),
            Output = ""
        };

        StringBuilder output = new StringBuilder();

        // Preview for next evolutions :
        //AltoXmlTextExporter altoXmlTextExporter = new AltoXmlTextExporter(
        //    NearestNeighbourWordExtractor.Instance,
        //    DocstrumBoundingBoxes.Instance);

        using (var document = PdfDocument.Open(data))
        {
            for (var i = 0; i < document.NumberOfPages; i++)
            {
                var page = document.GetPage(i + 1);

                var words = page.GetWords(NearestNeighbourWordExtractor.Instance);
                var images = page.GetImages();
                var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(words);

                var unsupervisedReadingOrderDetector = new UnsupervisedReadingOrderDetector(10);
                var orderedBlocks = unsupervisedReadingOrderDetector.Get(blocks);

                string? text = ContentOrderTextExtractor.GetText(page, true);
                output.AppendLine(text);

                // Extract images :
                if (options.ExtractImages)
                {
                    foreach (var image in images)
                    {
                        byte[] rawBytes = null;
                        string extension = "jpg";
                        if (image.TryGetPng(out rawBytes))
                            extension = "png";
                        else
                            rawBytes = image.RawBytes.ToArray();

                        var id = $"{Guid.NewGuid()}.{extension}";

                        var raw = $"![image](data:image/{extension};{id})";
                        output.AppendLine(raw);

                        result.Images.Add(new ImageRef()
                        {
                            Id = id,
                            Format = extension,
                            MarkdownRaw = raw,
                            RawBytes = rawBytes
                        });
                    }
                }
            }
        }

        // Convert file
        result.Output = output.ToString();

        return result;
    }

    #endregion
}
