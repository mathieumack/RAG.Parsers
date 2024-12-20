using System.Text;
using UglyToad.PdfPig.DocumentLayoutAnalysis.Export;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using System;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Collections.Generic;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using System.IO;
using RAG.Parsers.Pdf.Models;
using System.Linq;

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
    public ExtractOutput ToMarkdown(string filePath)
    {
        // Open file
        using var stream = File.OpenRead(filePath);

        // Convert file
        return ToMarkdown(stream);
    }

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    public ExtractOutput ToMarkdown(Stream data)
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

                string? text = ContentOrderTextExtractor.GetText(page);
                output.AppendLine(text);

                // Extract images :
                var images = page.GetImages();
                foreach(var image in images)
                {
                    result.Images.Add(new ImageRef()
                    {
                        Id = Guid.NewGuid().ToString(),
                        RawBytes = image.RawBytes
                    });
                    output.AppendLine($"![image](data:image/png;imageRefId,{result.Images.Last().Id})");
                }
            }
        }

        // Convert file
        result.Output = output.ToString();

        return result;
    }

    #endregion
}
