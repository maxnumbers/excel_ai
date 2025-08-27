using System;
using ExcelDna.Integration;
using System.Text;

namespace LocalAI.Excel.AddIn
{
    public static class ExcelHelper
    {
        public static string GetOptionalString(object value, string defaultValue)
        {
            if (value is ExcelMissing || value is ExcelEmpty || value == null)
                return defaultValue;
            return value.ToString();
        }

        public static double GetOptionalDouble(object value, double defaultValue)
        {
            if (value is ExcelMissing || value is ExcelEmpty || value == null)
                return defaultValue;
            if (double.TryParse(value.ToString(), out double result))
                return result;
            return defaultValue;
        }

        public static bool GetOptionalBool(object value, bool defaultValue)
        {
            if (value is ExcelMissing || value is ExcelEmpty || value == null)
                return defaultValue;
            if (bool.TryParse(value.ToString(), out bool result))
                return result;
            return defaultValue;
        }

        public static string ConvertRangeToText(object[,] range)
        {
            var sb = new StringBuilder();
            int rows = range.GetLength(0);
            int cols = range.GetLength(1);

            // Check if first row contains headers
            bool hasHeaders = true;
            for (int col = 0; col < cols; col++)
            {
                var value = range[0, col];
                if (value != null && double.TryParse(value.ToString(), out _))
                {
                    hasHeaders = false;
                    break;
                }
            }

            if (hasHeaders && rows > 1)
            {
                sb.Append("Headers: ");
                for (int col = 0; col < cols; col++)
                {
                    if (col > 0) sb.Append(", ");
                    sb.Append(range[0, col]?.ToString() ?? "");
                }
                sb.AppendLine();
                sb.AppendLine("Data rows:");

                for (int row = 1; row < Math.Min(rows, 11); row++)
                {
                    sb.Append($"Row {row}: ");
                    for (int col = 0; col < cols; col++)
                    {
                        if (col > 0) sb.Append(", ");
                        sb.Append(range[row, col]?.ToString() ?? "");
                    }
                    sb.AppendLine();
                }

                if (rows > 11)
                {
                    sb.AppendLine($"... and {rows - 11} more rows");
                }
            }
            else
            {
                for (int row = 0; row < Math.Min(rows, 10); row++)
                {
                    sb.Append($"Row {row + 1}: ");
                    for (int col = 0; col < cols; col++)
                    {
                        if (col > 0) sb.Append(", ");
                        sb.Append(range[row, col]?.ToString() ?? "");
                    }
                    sb.AppendLine();
                }

                if (rows > 10)
                {
                    sb.AppendLine($"... and {rows - 10} more rows");
                }
            }

            return sb.ToString();
        }
    }
}