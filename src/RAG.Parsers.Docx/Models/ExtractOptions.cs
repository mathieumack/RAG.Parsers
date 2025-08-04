namespace RAG.Parsers.Docx.Models;

public class ExtractOptions
{
    /// <summary>
    /// Extracts the images from the document
    /// </summary>
    /// <value>false</value>
    public bool ExtractImages { get; set; } = false;

    /// <summary>
    /// Extracts the tables from the document
    /// </summary>
    /// <value>true</value>
    public bool ExtractTables { get; set; } = true;

    /// <summary>
    /// Extracts the comments from the document
    /// </summary>
    /// <value>false</value>
    public bool ExtractComments { get; set; } = false;
}
