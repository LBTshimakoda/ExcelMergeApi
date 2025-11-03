using Microsoft.AspNetCore.Mvc;
using ExcelMergeApi.Models;
using ExcelMergeApi.Services;

namespace ExcelMergeApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelMergeController : ControllerBase
{
    private readonly ILogger<ExcelMergeController> _logger;
    private readonly IConfiguration _configuration;

    public ExcelMergeController(ILogger<ExcelMergeController> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
    }

    /// <summary>
    /// Merge multiple language-specific XLSX files into a single multilingual XLSX file
    /// </summary>
    /// <param name="zipFile">ZIP archive containing language folders with XLSX files</param>
    /// <param name="correlatedFile">Optional source language XLSX file</param>
    /// <param name="configJson">Optional JSON configuration string</param>
    [HttpPost("merge")]
    [RequestSizeLimit(500_000_000)]
    [RequestFormLimits(MultipartBodyLengthLimit = 500_000_000)]
    public async Task<IActionResult> MergeFiles([FromForm] MergeInput input)
    {
        if (input.ZipFile == null || input.ZipFile.Length == 0)
            return BadRequest(new { error = "ZIP file is required" });

        // Map input to configuration
        var defaultConfig = _configuration.GetSection("ExcelMerge").Get<ExcelMergeConfiguration>() ?? new ExcelMergeConfiguration();
        var config = new ExcelMergeConfiguration
        {
            SourceLanguage = input.SourceLanguage ?? defaultConfig.SourceLanguage,
            LanguageShifts = input.LanguageShifts ?? defaultConfig.LanguageShifts,
            LanguageCodes = input.LanguageCodes ?? defaultConfig.LanguageCodes,
            Tabs = input.Tabs ?? defaultConfig.Tabs,
            CopyColumns = input.CopyColumns ?? defaultConfig.CopyColumns,
            Filters = input.Filters ?? defaultConfig.Filters,
            ProjectSpecifics = input.ProjectSpecifics ?? defaultConfig.ProjectSpecifics
        };
        var service = new ExcelMergeService(config);

        using var zipStream = input.ZipFile.OpenReadStream();
        Stream? correlatedStream = input.CorrelatedFile?.OpenReadStream();

        var result = await service.MergeMultipleXlsxFilesIntoMultilingualFileAsync(
            zipStream,
            correlatedStream,
            input.ZipFile.FileName);

        if (!result.Success)
        {
            return BadRequest(new
            {
                error = result.Message,
                logs = result.Logs
            });
        }

        // Return only the merged file for now
        return File(
            result.MergedFileData!,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            result.MergedFileName ?? "merged.xlsx");
    }
    /// <summary>
    /// Get merge status and logs (for polling-based approach)
    /// </summary>
    [HttpGet("status/{jobId}")]
    public IActionResult GetStatus(string jobId)
    {
        // TODO: Implement job status tracking if using background processing
        return Ok(new { status = "Not implemented yet" });
    }

    /// <summary>
    /// Get current configuration
    /// </summary>
    [HttpGet("config")]
    public IActionResult GetConfiguration()
    {
        var config = LoadConfiguration(null);
        return Ok(config);
    }

    /// <summary>
    /// Load configuration from JSON string or appsettings
    /// </summary>
    private ExcelMergeConfiguration LoadConfiguration(string? configJson)
    {
        ExcelMergeConfiguration config;

        if (!string.IsNullOrEmpty(configJson))
        {
            try
            {
                config = System.Text.Json.JsonSerializer.Deserialize<ExcelMergeConfiguration>(configJson)
                    ?? new ExcelMergeConfiguration();
            }
            catch
            {
                config = new ExcelMergeConfiguration();
            }
        }
        else
        {
            // Load from appsettings.json
            config = _configuration.GetSection("ExcelMerge").Get<ExcelMergeConfiguration>()
                ?? new ExcelMergeConfiguration();
        }

        return config;
    }
}