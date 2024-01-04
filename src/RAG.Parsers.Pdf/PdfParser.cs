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
    public string ToMarkdown(string filePath)
    {
        StringBuilder output = new StringBuilder();

        // Preview for next evolutions :
        //AltoXmlTextExporter altoXmlTextExporter = new AltoXmlTextExporter(
        //    NearestNeighbourWordExtractor.Instance,
        //    DocstrumBoundingBoxes.Instance);

        using (var document = PdfDocument.Open(filePath))
        {
            for (var i = 0; i < document.NumberOfPages; i++)
            {
                var page = document.GetPage(i + 1);

                string? text = ContentOrderTextExtractor.GetText(page);
                output.AppendLine(text);
            }
        }

        // Convert file
        return output.ToString();
    }

    #endregion
}
