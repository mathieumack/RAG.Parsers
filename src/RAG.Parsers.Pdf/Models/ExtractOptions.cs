namespace RAG.Parsers.Pdf.Models;

public class ExtractOptions
{
    /// <summary>
    /// Extracts the images from the document
    /// </summary>
    /// <value>false</value>
    public bool ExtractImages { get; set; } = false;

    /// <summary>
    /// Extract the pages as images
    /// </summary>
    public bool ExtractPageImages { get; set; } = false;

    /// <summary>
    /// Extracts the tables from the document
    /// </summary>
    /// <value>true</value>
    public bool ExtractTables { get; set; } = true;
}
