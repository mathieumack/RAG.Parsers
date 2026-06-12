namespace RAG.Parsers.Msg.Models;

public class ExtractOptions
{
    /// <summary>
    /// Extracts attachments from the message and surfaces them as references.
    /// </summary>
    /// <value>true</value>
    public bool ExtractAttachments { get; set; } = true;
}
