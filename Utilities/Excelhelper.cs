using ExcelMergeApi.Models;

namespace ExcelMergeApi.Utilities;

public static class ExcelHelper
{
    /// <summary>
    /// Convert column name (A, B, AA, etc.) to column number (1, 2, 27, etc.)
    /// </summary>
    public static int ColumnNameToNumber(string columnName)
    {
        if (string.IsNullOrEmpty(columnName))
            throw new ArgumentException("Column name cannot be null or empty");

        columnName = columnName.ToUpper();
        int result = 0;

        for (int i = 0; i < columnName.Length; i++)
        {
            result *= 26;
            result += (columnName[i] - 'A' + 1);
        }

        return result;
    }

    /// <summary>
    /// Convert column number (1, 2, 27, etc.) to column name (A, B, AA, etc.)
    /// </summary>
    public static string ColumnNumberToName(int columnNumber)
    {
        if (columnNumber <= 0)
            throw new ArgumentException("Column number must be greater than 0");

        string columnName = "";

        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// Get language code from path in ZIP archive
    /// </summary>
    public static string GetLanguageCodeFromPathInZip(string path)
    {
        if (string.IsNullOrEmpty(path))
            return "";

        int slashIndex = path.IndexOf('/');
        int backslashIndex = path.IndexOf('\\');

        if (slashIndex != -1)
            return path.Substring(0, slashIndex);
        else if (backslashIndex != -1)
            return path.Substring(0, backslashIndex);
        else
            return path;
    }

    /// <summary>
    /// Check if a value passes a filter condition
    /// </summary>
    public static bool CheckFilter(string fvalue, FilterConfiguration filter)
    {
        if (filter.Type.Equals("string", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrEmpty(fvalue))
                fvalue = "empty";

            if (filter.Condition == "=" && filter.Value == fvalue)
                return false;
            if (filter.Condition == "!=" && filter.Value != fvalue)
                return false;
        }
        else if (filter.Type.Equals("digit", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrEmpty(fvalue))
                return true;

            if (!int.TryParse(filter.Value, out int fval0))
                return true;

            if (!int.TryParse(fvalue, out int fval))
                fval = 1;

            return filter.Condition switch
            {
                "=" => fval != fval0,
                "!=" => fval == fval0,
                "<=" => fval > fval0,
                ">=" => fval < fval0,
                "<" => fval >= fval0,
                ">" => fval <= fval0,
                _ => true
            };
        }

        return true;
    }
}