using System.Text;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using System;
using System.IO;
using RAG.Parsers.Pdf.Models;
using System.Linq;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;
using System.Collections.Generic;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.Rendering.Skia;
using UglyToad.PdfPig.Graphics.Colors;
using Microsoft.Extensions.Logging;
using static SkiaSharp.HarfBuzz.SKShaper;
using System.Reflection.Metadata;

namespace RAG.Parsers.Pdf;

/// <summary>
/// Pdf class for read document
/// </summary>
public class PdfParser : IDisposable
{
    #region Constants

    private const double VERTICAL_THRESHOLD = 80;

    #endregion

    private readonly ILogger<PdfParser> logger;
    private Stream sourceStream;
    private PdfDocument pdfDocument;
    private readonly Dictionary<string, IPdfImage> imageRefs = new();
    private readonly Dictionary<string, PageRef> pageRefs = new();

    public PdfParser(ILogger<PdfParser> logger)
    {
        this.logger = logger;
    }

    #region Public Methods

    /// <summary>
    /// Read file and open it
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public ExtractOutput ToMarkdown(string filePath, ExtractOptions options)
    {
        // Open file
        sourceStream = File.OpenRead(filePath);

        // Convert file
        return ToMarkdown(sourceStream, options);
    }

    /// <summary>
    /// Converts a PDF stream to Markdown format
    /// </summary>
    /// <param name="data">PDF data stream</param>
    /// <param name="options">Extraction options</param>
    /// <returns>Extracted content including text and images</returns>
    public ExtractOutput ToMarkdown(Stream data, ExtractOptions options)
    {
        var result = InitializeExtractOutput();
        var output = new StringBuilder();

        // Clear old reference to old document
        if (sourceStream != null)
        {
            sourceStream.Dispose();
            sourceStream = null;
        }
        if (pdfDocument != null)
        {
            pdfDocument.Dispose();
            pdfDocument = null;
        }

        imageRefs.Clear();
        pageRefs.Clear();

        pdfDocument = PdfDocument.Open(data);
        if(options.ExtractPageImages)
            pdfDocument.AddSkiaPageFactory();

        ProcessDocument(pdfDocument, options, output, result);

        result.Output = output.ToString();

        return result;
    }

    /// <summary>
    /// Initializes a new instance of the ExtractOutput class.
    /// </summary>
    private ExtractOutput InitializeExtractOutput()
    {
        return new ExtractOutput
        {
            Images = new List<ImageRef>(),
            Pages = new List<PageRef>(),
            Output = string.Empty
        };
    }

    /// <summary>
    /// Processes the PDF document and extracts content based on the provided options.
    /// </summary>
    /// <param name="document">The PDF document to process.</param>
    /// <param name="options">The extraction options.</param>
    /// <param name="output">The StringBuilder to store the extracted content.</param>
    /// <param name="result">The ExtractOutput object to store the result.</param>
    private void ProcessDocument(PdfDocument document, ExtractOptions options, StringBuilder output, ExtractOutput result)
    {
        for (var i = 0; i < document.NumberOfPages; i++)
        {
            var page = document.GetPage(i + 1);

            ProcessPage(document, page, options, output, result);
        }
    }

    /// <summary>
    /// Processes a single page of the PDF document and extracts content based on the provided options.
    /// </summary>
    /// <param name="page">The page to process.</param>
    /// <param name="options">The extraction options.</param>
    /// <param name="output">The StringBuilder to store the extracted content.</param>
    /// <param name="result">The ExtractOutput object to store the result.</param>
    private void ProcessPage(PdfDocument document, Page page, ExtractOptions options, StringBuilder output, ExtractOutput result)
    {
        var words = page.GetWords(NearestNeighbourWordExtractor.Instance);
        var images = options.ExtractImages ? page.GetImages().ToList() : new();
        var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(words);

        var orderedBlocks = new UnsupervisedReadingOrderDetector(10).Get(blocks);
        var currentPageIndex = output.Length == 0 ? 0 : output.Length - 1;

        ProcessBlocks(orderedBlocks, images, output, result);
        ProcessRemainingImages(images, page, currentPageIndex, output, result);
        ProcessPage(document, page.Number, options, output, result);
    }

    private void ProcessPage(PdfDocument document, int pageNumber, ExtractOptions options, StringBuilder output, ExtractOutput result)
    {
        if (options.ExtractPageImages)
        {
            var pageRef = CreatePageRef(pageNumber, "png");
            result.Pages.Add(pageRef);
            pageRefs.Add(pageRef.Id, pageRef);

            //output.AppendLine($"Page {pageNumber} image view :");
            output.AppendLine(pageRef.MarkdownRaw);
        }
    }

    /// <summary>
    /// Creates an image reference from the raw bytes and extension of an image.
    /// </summary>
    /// <param name="pageNumber"></param>
    /// <param name="rawBytes">The raw bytes of the image.</param>
    /// <param name="extension">The extension of the image.</param>
    /// <returns>An ImageRef object containing the image reference.</returns>
    private PageRef CreatePageRef(int pageNumber, string extension)
    {
        var id = $"{Guid.NewGuid()}.{extension}";
        var raw = $"![image](data:image/{extension};{id})";

        return new PageRef
        {
            Id = id,
            Format = extension,
            MarkdownRaw = raw,
            PageNumber = pageNumber
        };
    }

    /// <summary>
    /// Processes the text blocks and images on a page and extracts content based on the provided options.
    /// </summary>
    /// <param name="orderedBlocks">The ordered text blocks to process.</param>
    /// <param name="images">The list of images on the page.</param>
    /// <param name="output">The StringBuilder to store the extracted content.</param>
    /// <param name="result">The ExtractOutput object to store the result.</param>
    private void ProcessBlocks(IEnumerable<TextBlock> orderedBlocks, List<IPdfImage> images, StringBuilder output, ExtractOutput result)
    {
        foreach (var block in orderedBlocks)
        {
            var nearImages = GetNearbyImages(block, images);

            ProcessBlockWithImages(block, nearImages, output, result);
        }
    }

    /// <summary>
    /// Processes a text block and its nearby images, and extracts content based on the provided options.
    /// </summary>
    /// <param name="block">The text block to process.</param>
    /// <param name="nearImages">The dictionary of nearby images categorized by their position.</param>
    /// <param name="output">The StringBuilder to store the extracted content.</param>
    /// <param name="result">The ExtractOutput object to store the result.</param>
    private void ProcessBlockWithImages(TextBlock block, Dictionary<string, List<IPdfImage>> nearImages, StringBuilder output, ExtractOutput result)
    {
        if (nearImages["top"].Count + nearImages["bottom"].Count > 0)
        {
            ProcessImagesForPosition(nearImages["top"], output, result);

            output.AppendLine(block.Text);

            ProcessImagesForPosition(nearImages["bottom"], output, result);
        }
        else
        {
            output.AppendLine(block.Text);
        }
    }

    /// <summary>
    /// Processes a list of images for a specific position and extracts content based on the provided options.
    /// </summary>
    /// <param name="images">The list of images to process.</param>
    /// <param name="output">The StringBuilder to store the extracted content.</param>
    /// <param name="result">The ExtractOutput object to store the result.</param>
    private void ProcessImagesForPosition(List<IPdfImage> images, StringBuilder output, ExtractOutput result)
    {
        foreach (var image in images)
        {
            var imageRef = CreateImageReference(image);

            output.AppendLine(imageRef.MarkdownRaw);
            result.Images.Add(imageRef);
        }
    }

    /// <summary>
    /// Processes the remaining images on a page and extracts content based on the provided options.
    /// </summary>
    /// <param name="images">The list of remaining images to process.</param>
    /// <param name="page">The page containing the images.</param>
    /// <param name="currentPageIndex">The current index in the output StringBuilder.</param>
    /// <param name="output">The StringBuilder to store the extracted content.</param>
    /// <param name="result">The ExtractOutput object to store the result.</param>
    private void ProcessRemainingImages(List<IPdfImage> images, Page page, int currentPageIndex, StringBuilder output, ExtractOutput result)
    {
        foreach (var image in images)
        {
            var imageRef = CreateImageReference(image);

            if (image.Bounds.Bottom > page.Height / 2)
            {
                output.Insert(currentPageIndex, imageRef.MarkdownRaw + Environment.NewLine);
            }
            else
            {
                output.AppendLine(imageRef.MarkdownRaw);
            }

            result.Images.Add(imageRef);
        }
    }

    /// <summary>
    /// Gets the raw bytes and extension of an image.
    /// </summary>
    /// <param name="id">The image id to process based on ImageRef object.</param>
    /// <returns>A tuple containing the raw bytes and extension of the image.</returns>
    public EntryRefContent GetImageBytesAndExtension(string id)
    {
        if (id is null || !imageRefs.ContainsKey(id))
            return null;

        byte[] rawBytes = null;
        string extension = "jpg";

        var image = imageRefs[id];
        if (image.TryGetPng(out rawBytes))
        {
            extension = "png";
        }
        else
        {
            rawBytes = image.RawBytes.ToArray();
        }

        return new EntryRefContent()
        {
            Extension = extension,
            RawBytes = rawBytes
        };
    }

    /// <summary>
    /// Gets a 
    /// </summary>
    /// <param name="id"></param>
    /// <returns></returns>
    public EntryRefContent GetPageImageBytesandExtension(string id)
    {
        if(pdfDocument is null && string.IsNullOrWhiteSpace(id) && !pageRefs.ContainsKey(id))
            return null;

        var page = pageRefs[id];
        using (var ms = pdfDocument.GetPageAsPng(page.PageNumber, 1, RGBColor.White))
        {
            return new EntryRefContent()
            {
                Extension = "png",
                RawBytes = ms.ToArray()
            };
        }
    }

    /// <summary>
    /// Creates an image reference from the raw bytes and extension of an image.
    /// </summary>
    /// <param name="rawBytes">The raw bytes of the image.</param>
    /// <param name="extension">The extension of the image.</param>
    /// <returns>An ImageRef object containing the image reference.</returns>
    private ImageRef CreateImageReference(IPdfImage image)
    {
        var id = Guid.NewGuid().ToString();
        var raw = $"![image](data:image/jpg;{id})";

        var result = new ImageRef
        {
            Id = id,
            MarkdownRaw = raw
        };

        imageRefs.Add(result.Id, image);

        return result;
    }

    /// <summary>
    /// Gets the nearby images for a text block.
    /// </summary>
    /// <param name="block">The text block to process.</param>
    /// <param name="images">The list of images on the page.</param>
    /// <returns>A dictionary of nearby images categorized by their position.</returns>
    private Dictionary<string, List<IPdfImage>> GetNearbyImages(TextBlock block, List<IPdfImage> images)
    {
        var nearImages = new Dictionary<string, List<IPdfImage>>
                         {
                             { "top", new List<IPdfImage>() },
                             { "bottom", new List<IPdfImage>() }
                         };

        var imagesToRemove = new List<IPdfImage>();
        foreach (var image in images.OrderBy(img => img.Bounds.TopLeft.Y))
        {
            if (IsImageAboveBlock(image, block))
            {
                nearImages["top"].Add(image);

                imagesToRemove.Add(image);
            }
            else if (IsImageNearBlock(image, block))
            {
                nearImages["bottom"].Add(image);

                imagesToRemove.Add(image);
            }
        }

        foreach (var image in imagesToRemove)
        {
            images.Remove(image);
        }

        return nearImages;
    }

    /// <summary>
    /// Determines if an image is above a text block.
    /// </summary>
    /// <param name="image">The image to check.</param>
    /// <param name="block">The text block to check against.</param>
    /// <returns>True if the image is above the text block, otherwise false.</returns>
    private bool IsImageAboveBlock(IPdfImage image, TextBlock block)
    {
        return image.Bounds.BottomLeft.Y >= block.BoundingBox.TopLeft.Y &&
               Math.Abs(image.Bounds.BottomLeft.Y - block.BoundingBox.TopLeft.Y) < VERTICAL_THRESHOLD;
    }

    /// <summary>
    /// Determines if an image is near a text block, depending on a threshold value.
    /// </summary>
    /// <param name="image">The image to check.</param>
    /// <param name="block">The text block to check against.</param>
    /// <returns>True if the image is near the text block, otherwise false.</returns>
    private bool IsImageNearBlock(IPdfImage image, TextBlock block)
    {
        var hasVerticalOverlap = Math.Abs(image.Bounds.TopLeft.Y - block.BoundingBox.TopLeft.Y) <= VERTICAL_THRESHOLD ||
                                 Math.Abs(image.Bounds.BottomLeft.Y - block.BoundingBox.BottomLeft.Y) <= VERTICAL_THRESHOLD;

        var isBelow = image.Bounds.TopLeft.Y <= block.BoundingBox.BottomLeft.Y &&
                      Math.Abs(image.Bounds.TopLeft.Y - block.BoundingBox.BottomLeft.Y) < VERTICAL_THRESHOLD;
        var isLeft = image.Bounds.BottomRight.X <= block.BoundingBox.BottomLeft.X && hasVerticalOverlap;
        var isRight = image.Bounds.BottomLeft.X >= block.BoundingBox.BottomRight.X && hasVerticalOverlap;

        return isBelow || isLeft || isRight;
    }

    public void Dispose()
    {
        if (sourceStream != null)
        {
            sourceStream.Dispose();
            sourceStream = null;
        }
        if (pdfDocument != null)
        {
            pdfDocument.Dispose();
            pdfDocument = null;
        }

        // Suppress finalization to adhere to CA1816  
        GC.SuppressFinalize(this);
    }

    #endregion
}
