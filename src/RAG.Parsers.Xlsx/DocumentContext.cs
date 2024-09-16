namespace RAG.Parsers.Xlsx;

internal class DocumentContext
{
    #region Properties

    public const string DefaultSheetNumberTemplate = "\n# Worksheet \"{name}\"\n";
    public const char DefaultCellBalise = '|';

    public bool WithQuotes { get; set; } = false;

    public string WorksheetNumberTemplate { get; set; } = DefaultSheetNumberTemplate;

    #endregion
}
