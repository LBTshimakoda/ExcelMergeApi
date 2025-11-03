// Models/MergeInput.cs
using ExcelMergeApi.Models;

public class MergeInput
{
    public IFormFile ZipFile { get; set; } = default!;
    public IFormFile? CorrelatedFile { get; set; }
    // Expose config properties directly, or use a nested config object
    public string? SourceLanguage { get; set; }
    public Dictionary<string, int>? LanguageShifts { get; set; }
    public Dictionary<string, string>? LanguageCodes { get; set; }
    public Dictionary<string, string>? Tabs { get; set; }
    public List<string>? CopyColumns { get; set; }
    public List<FilterConfiguration>? Filters { get; set; }
    public string? ProjectSpecifics { get; set; }
}