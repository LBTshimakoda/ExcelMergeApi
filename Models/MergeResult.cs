// Models/MergeResult.cs
public class MergeResult
{
    public byte[] MergedFileData { get; set; } = default!;
    public string? MergedFileName { get; set; }
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public List<string> Logs { get; set; } = new();
}
