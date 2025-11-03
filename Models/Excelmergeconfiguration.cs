namespace ExcelMergeApi.Models;

public class ExcelMergeConfiguration
{
    public string SourceLanguage { get; set; } = "en-gb";
    public Dictionary<string, int> LanguageShifts { get; set; } = new()
    {
        {"en-gb", 0},
        {"it-it", 1},
        {"fr-fr", 2},
        {"de-de", 3},
        {"es-es", 4},
        {"en-us", 5},
        {"en-xn", 6},
        {"ru-ru", 7},
        {"ja-jp", 8},
        {"ar-xm", 9},
        {"zh-cn", 10},
        {"ko-kr", 11}
    };

    public int lang_row { get; set; } = 1;

    public Dictionary<string, string> LanguageCodes { get; set; } = new();

    public Dictionary<string, string> Tabs { get; set; } = new();

    public List<string> CopyColumns { get; set; } = new() { };

    public List<FilterConfiguration> Filters { get; set; } = new();

    public string ProjectSpecifics { get; set; } = "";
}

public class FilterConfiguration
{
    public string Column { get; set; } = "";
    public string Type { get; set; } = "string"; // "string" or "digit"
    public string Condition { get; set; } = "="; // =, !=, <, >, <=, >=
    public string Value { get; set; } = "";
}