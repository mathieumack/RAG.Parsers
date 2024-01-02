using System.Text;
using UglyToad.PdfPig.DocumentLayoutAnalysis.Export;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using System;

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
        throw new NotImplementedException("This feature is not available yet");

        StringBuilder output = new StringBuilder();

        AltoXmlTextExporter altoXmlTextExporter = new AltoXmlTextExporter(
            NearestNeighbourWordExtractor.Instance,
            DocstrumBoundingBoxes.Instance);

        using (var document = PdfDocument.Open(filePath))
        {
            for (var i = 0; i < document.NumberOfPages; i++)
            {
                var page = document.GetPage(i + 1);

                // Convert page to text, adding a reference to hocrjs script
                string xml = altoXmlTextExporter.Get(page);
                
                // Save text to an html file
                output.AppendLine(xml);
            }
        }

        // Convert file
        return output.ToString();
    }

    #endregion
}
