using System.IO.Compression;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using ExcelMergeApi.Models;
using ExcelMergeApi.Utilities;

namespace ExcelMergeApi.Services;

public class ExcelMergeService
{
    private readonly List<string> _logs = new();
    private readonly ExcelMergeConfiguration _config;

    public ExcelMergeService(ExcelMergeConfiguration config)
    {
        _config = config;
        // Set EPPlus license context for non-commercial use
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public List<string> GetLogs() => _logs;

    private void AddLog(string message)
    {
        _logs.Add($"{DateTime.Now}: {message}");
    }

    /// <summary>
    /// Main entry point for merging multiple XLSX files into a multilingual file
    /// </summary>
    public async Task<MergeResult> MergeMultipleXlsxFilesIntoMultilingualFileAsync(
        Stream zipFileStream,
        Stream? correlatedFileStream,
        string zipFileName)
    {
        try
        {
            AddLog("Starting merge process");
            AddLog("Checking configuration");

            // Create backup handling
            var inputFile = zipFileName;

            // Get XLSX files from delivery kit
            var languageBasedXlsxList = await GetXlsxFilesFromDeliveryKitAsync(zipFileStream, correlatedFileStream);

            AddLog("Language files collected:");
            var langListToCheck = new List<string>();

            foreach (var item in languageBasedXlsxList)
            {
                AddLog($"{item.TMSLanguageCode} {item.Filename}");
                if (!langListToCheck.Contains(item.TMSLanguageCode))
                    langListToCheck.Add(item.TMSLanguageCode);
            }

            // Validate source language files
            if (!langListToCheck.Contains(_config.SourceLanguage))
            {
                var errorMsg = $"Source language files not available. Please ensure source files exist for: {_config.SourceLanguage}";
                AddLog(errorMsg);
                return new MergeResult
                {
                    Success = false,
                    Message = errorMsg,
                    Logs = _logs
                };
            }

            // Get source XLSX files
            var sourceBaseXlsxList = GetBaseXlsxFiles(languageBasedXlsxList);
            AddLog($"Base file selected: {sourceBaseXlsxList[0].TMSLanguageCode}");

            // Build multilingual XLSX
            foreach (var singleSourceLanguageBasedXlsx in sourceBaseXlsxList)
            {
                if (_config.Tabs != null && _config.Tabs.Count > 0)
                {
                    singleSourceLanguageBasedXlsx.XlsxDocumentStream =
                        await BuildMultilingualXlsxWithTabsAsync(singleSourceLanguageBasedXlsx, languageBasedXlsxList);
                }
                else
                {
                    singleSourceLanguageBasedXlsx.XlsxDocumentStream =
                        await BuildMultilingualXlsxAsync(singleSourceLanguageBasedXlsx, languageBasedXlsxList);
                }
                AddLog("Languages merged");
            }

            // Get the merged file data
            var mergedFileData = sourceBaseXlsxList[0].XlsxDocumentStream.ToArray();

            AddLog("Merge completed successfully");

            return new MergeResult
            {
                Success = true,
                Message = "Merge completed successfully",
                Logs = _logs,
                MergedFileData = mergedFileData,
                MergedFileName = sourceBaseXlsxList[0].Filename
            };
        }
        catch (Exception ex)
        {
            AddLog($"Error during merge: {ex.Message}");
            return new MergeResult
            {
                Success = false,
                Message = $"Error: {ex.Message}",
                Logs = _logs
            };
        }
    }

    /// <summary>
    /// Extract XLSX files from ZIP archive
    /// </summary>
    private async Task<List<LanguageBasedXlsx>> GetXlsxFilesFromDeliveryKitAsync(
        Stream zipFileStream,
        Stream? correlatedFileStream)
    {
        var languageBasedXlsxList = new List<LanguageBasedXlsx>();
        var fileslist = new List<string>();

        try
        {
            using (var zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Read, true))
            {
                foreach (var file in zipArchive.Entries)
                {
                    if (file.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        using (var stream = file.Open())
                        {
                            var temporaryStream = new MemoryStream();
                            await stream.CopyToAsync(temporaryStream);
                            temporaryStream.Position = 0;

                            // Validate XLSX file
                            ValidateXlsxFile(temporaryStream, file.FullName);

                            var languageBasedXlsx = new LanguageBasedXlsx
                            {
                                Path = file.FullName,
                                Filename = file.Name,
                                TMSLanguageCode = ExcelHelper.GetLanguageCodeFromPathInZip(file.FullName).ToLower(),
                                XlsxDocumentStream = temporaryStream
                            };

                            languageBasedXlsxList.Add(languageBasedXlsx);

                            if (!fileslist.Contains(file.Name))
                                fileslist.Add(file.Name);
                        }
                    }
                }
            }

            // Handle correlated file if provided
            if (correlatedFileStream != null)
            {
                foreach (var fileName in fileslist)
                {
                    var temporaryStream = new MemoryStream();
                    correlatedFileStream.Position = 0;
                    await correlatedFileStream.CopyToAsync(temporaryStream);
                    temporaryStream.Position = 0;

                    ValidateXlsxFile(temporaryStream, fileName);

                    var languageBasedXlsx = new LanguageBasedXlsx
                    {
                        Path = _config.SourceLanguage + "/" + fileName,
                        Filename = fileName,
                        TMSLanguageCode = _config.SourceLanguage,
                        XlsxDocumentStream = temporaryStream
                    };

                    languageBasedXlsxList.Add(languageBasedXlsx);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"GetXlsxFilesFromDeliveryKit: {ex.Message}", ex);
        }

        return languageBasedXlsxList;
    }

    /// <summary>
    /// Validate that file is a valid XLSX
    /// </summary>
    private void ValidateXlsxFile(MemoryStream xlsxMemoryStream, string fullFilePath)
    {
        try
        {
            xlsxMemoryStream.Position = 0;
            using (var package = new ExcelPackage(xlsxMemoryStream))
            {
                // Just opening the package validates it
            }
            xlsxMemoryStream.Position = 0;
        }
        catch (Exception ex)
        {
            throw new Exception($"Could not load file {fullFilePath} as an Excel Package. File might be corrupted. Ex: {ex.Message}");
        }
    }

    /// <summary>
    /// Get source language XLSX files to use as base
    /// </summary>
    private List<SourceLanguageBasedXlsx> GetBaseXlsxFiles(List<LanguageBasedXlsx> languageBasedXlsxList)
    {
        var sourceLanguageBasedXlsxList = new List<SourceLanguageBasedXlsx>();

        var groupedByLanguageList = languageBasedXlsxList
            .OrderBy(n => n.Path.Substring(0, Math.Max(n.Path.IndexOf('/'), n.Path.IndexOf('\\'))))
            .GroupBy(n => n.Path.Substring(0, Math.Max(n.Path.IndexOf('/'), n.Path.IndexOf('\\'))));

        foreach (var item in groupedByLanguageList)
        {
            var first = item.FirstOrDefault();
            if (first != null && ExcelHelper.GetLanguageCodeFromPathInZip(first.Path) == _config.SourceLanguage)
            {
                sourceLanguageBasedXlsxList.Add(new SourceLanguageBasedXlsx(
                    first.XlsxDocumentStream,
                    first.Filename,
                    first.Path,
                    ExcelHelper.GetLanguageCodeFromPathInZip(first.Path)));
            }
        }

        return sourceLanguageBasedXlsxList;
    }

    /// <summary>
    /// Build multilingual XLSX by merging translations
    /// </summary>
    private async Task<MemoryStream> BuildMultilingualXlsxAsync(
        SourceLanguageBasedXlsx sourceBaseXlsx,
        List<LanguageBasedXlsx> languageBasedXlsxList)
    {
        var sSourceFileName = sourceBaseXlsx.Filename;

        sourceBaseXlsx.XlsxDocumentStream.Position = 0;
        using (var sourcePackage = new ExcelPackage(sourceBaseXlsx.XlsxDocumentStream))
        {
            var sourceWorksheet = sourcePackage.Workbook.Worksheets.FirstOrDefault();
            if (sourceWorksheet == null)
                throw new Exception("No worksheet found in source file");

            var dimension = sourceWorksheet.Dimension;
            if (dimension == null)
                throw new Exception("Worksheet has no data");

            int rowsCount = dimension.End.Row;
            int columnsCount = dimension.End.Column;

            // Find language row
            int langRow = _config.lang_row;
            //for (int i = 1; i <= 3; i++)
            //{
            //    var value = sourceWorksheet.Cells[i, ExcelHelper.ColumnNameToNumber(_config.CopyColumns[0])].Text;
            //    if (!string.IsNullOrEmpty(value) && Regex.IsMatch(value, @"^en\s?"))
            //    {
            //        langRow = i;
            //        break;
            //    }
            //}

            // Find max row with data
            int maxRow = 0;
            for (int i = 2; i <= rowsCount; i++)
            {
                var value = sourceWorksheet.Cells[i, ExcelHelper.ColumnNameToNumber(_config.CopyColumns[0])].Text;
                if (!string.IsNullOrEmpty(value))
                    maxRow = i;
            }
            rowsCount = maxRow;

            // Process each language file
            foreach (var singleLanguageBasedXlsx in languageBasedXlsxList)
            {
                if (sSourceFileName != singleLanguageBasedXlsx.Filename)
                    continue;

                if (sourceBaseXlsx.TMSLanguageCode == singleLanguageBasedXlsx.TMSLanguageCode)
                    continue;

                if (!_config.LanguageShifts.ContainsKey(singleLanguageBasedXlsx.TMSLanguageCode))
                    continue;

                int langShift = _config.LanguageShifts[singleLanguageBasedXlsx.TMSLanguageCode];

                singleLanguageBasedXlsx.XlsxDocumentStream.Position = 0;
                using (var targetPackage = new ExcelPackage(singleLanguageBasedXlsx.XlsxDocumentStream))
                {
                    var targetWorksheet = targetPackage.Workbook.Worksheets.FirstOrDefault();
                    if (targetWorksheet == null)
                        continue;

                    // Copy columns
                    foreach (var col in _config.CopyColumns)
                    {
                        int colNum = ExcelHelper.ColumnNameToNumber(col);

                        for (int i = langRow + 1; i <= rowsCount; i++)
                        {
                            var value = targetWorksheet.Cells[i, colNum].Text;
                            var sourceValue = sourceWorksheet.Cells[i, colNum].Text;

                            // Apply filters
                            bool rowFilter = true;
                            foreach (var filter in _config.Filters)
                            {
                                var fvalue = targetWorksheet.Cells[i, ExcelHelper.ColumnNameToNumber(filter.Column)].Text;
                                if (!string.IsNullOrEmpty(fvalue))
                                {
                                    rowFilter = rowFilter && ExcelHelper.CheckFilter(fvalue, filter);
                                }
                            }

                            if (rowFilter || singleLanguageBasedXlsx.TMSLanguageCode == _config.SourceLanguage)
                            {
                                if (!string.IsNullOrEmpty(sourceValue) && string.IsNullOrEmpty(value))
                                {
                                    AddLog($"Empty value found for language {singleLanguageBasedXlsx.TMSLanguageCode} in cell {col}{i}");
                                }

                                if (!string.IsNullOrEmpty(value))
                                {
                                    sourceWorksheet.Cells[i, colNum + langShift].Value = value;

                                    // Handle project specifics (Versace)
                                    if (_config.ProjectSpecifics.Equals("versace product descriptions", StringComparison.OrdinalIgnoreCase)
                                        && col == _config.CopyColumns[0])
                                    {
                                        var targetCol = ExcelHelper.ColumnNumberToName(colNum + langShift);
                                        sourceWorksheet.Cells[i, colNum + langShift + 1].Formula = $"=LEN({targetCol}{i})";
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Save to memory stream
            var finalStream = new MemoryStream();
            await sourcePackage.SaveAsAsync(finalStream);
            finalStream.Position = 0;
            return finalStream;
        }
    }

    /// <summary>
    /// Build multilingual XLSX with tabs support (placeholder for future implementation)
    /// </summary>
    private async Task<MemoryStream> BuildMultilingualXlsxWithTabsAsync(
        SourceLanguageBasedXlsx sourceBaseXlsx,
        List<LanguageBasedXlsx> languageBasedXlsxList)
    {
        // This would be similar to BuildMultilingualXlsx but with tab-specific logic
        // For now, delegate to the main method
        return await BuildMultilingualXlsxAsync(sourceBaseXlsx, languageBasedXlsxList);
    }
}