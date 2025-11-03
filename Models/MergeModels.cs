namespace ExcelMergeApi.Models;

public class LanguageBasedXlsx
{
    public string Path { get; set; } = "";
    public string Filename { get; set; } = "";
    public string TMSLanguageCode { get; set; } = "";
    public MemoryStream XlsxDocumentStream { get; set; } = new();
}

public class SourceLanguageBasedXlsx : LanguageBasedXlsx
{
    public SourceLanguageBasedXlsx() { }

    public SourceLanguageBasedXlsx(MemoryStream stream, string filename, string path, string languageCode)
    {
        XlsxDocumentStream = stream;
        Filename = filename;
        Path = path;
        TMSLanguageCode = languageCode;
    }
}

public class MergeRequest
{
    public IFormFile ZipFile { get; set; } = null!;
    public IFormFile? CorrelatedFile { get; set; }
    //public ExcelMergeConfiguration? Configuration { get; set; }
    public string? ConfigJson { get; set; } // User pastes JSON here
    // Flattened ExcelMergeConfiguration properties
    //public string? SourceLanguage { get; set; }
    //public Dictionary<string, int>? LanguageShifts { get; set; }
    //public Dictionary<string, string>? LanguageCodes { get; set; }
    //public Dictionary<string, string>? Tabs { get; set; }
    //public List<string>? CopyColumns { get; set; }
    //public List<FilterConfiguration>? Filters { get; set; }
    //public string? ProjectSpecifics { get; set; }
}

public class MergeProgressUpdate
{
    public string Message { get; set; } = "";
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public int PercentComplete { get; set; } = 0;
}

