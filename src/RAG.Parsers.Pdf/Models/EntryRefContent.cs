using System.Collections.Generic;

namespace RAG.Parsers.Pdf.Models;

public class EntryRefContent
{
    /// <summary>
    /// Image format extension
    /// </summary>
    public string Extension { get; set; }

    /// <summary>
    /// Image memory byte array
    /// </summary>
    public byte[] RawBytes { get; set; }
}
