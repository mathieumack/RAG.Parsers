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
using static System.Net.Mime.MediaTypeNames;

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
                var images = options.ExtractImages ? page.GetImages().ToList() : new();
                var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(words);

                var unsupervisedReadingOrderDetector = new UnsupervisedReadingOrderDetector(10);
                var orderedBlocks = unsupervisedReadingOrderDetector.Get(blocks);

                // Now for each bloc, we can extract the text and link images that are really next to it :
                foreach (var block in orderedBlocks)
                {
                    output.AppendLine(block.Text);
                    // Search for images next to the block :
                    // Images coordonnates :
                    //image.Bounds.TopLeft.X;
                    //image.Bounds.TopLeft.Y;
                    //image.Bounds.TopRight.X;
                    //image.Bounds.TopRight.Y;

                    // text block coordonnates :
                    //block.BoundingBox.TopLeft.X;
                    //block.BoundingBox.TopLeft.Y;
                    //block.BoundingBox.TopRight.X;
                    //block.BoundingBox.TopRight.Y;

                    // Find image that is near text block :
                    var nearImage = images.FirstOrDefault(x =>
                    {
                        var imageX = x.Bounds.TopLeft.X;
                        var imageY = x.Bounds.TopLeft.Y;
                        var blockX = block.BoundingBox.TopLeft.X;
                        var blockY = block.BoundingBox.TopLeft.Y;
                        return Math.Abs(imageX - blockX) < 100 && Math.Abs(imageY - blockY) < 100;
                    });
                    if (nearImage != null)
                    {
                        byte[] rawBytes = null;
                        string extension = "jpg";
                        if (nearImage.TryGetPng(out rawBytes))
                            extension = "png";
                        else
                            rawBytes = nearImage.RawBytes.ToArray();
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

                        // Delete image from list to avoid duplicates :
                        images.Remove(nearImage);
                    }
                }

                // Now add other images that are not near text :
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

        // Convert file
        result.Output = output.ToString();

        return result;
    }

    #endregion
}
