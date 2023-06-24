using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace ExcelDataMerger
{
    public class Tester
    {
        public void UpdateColumns(string sourceFolderPath, string destinationFolderPath, string sourceNames, string sourceValues, string destinationNames, string destinationValues)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var destinationDirectory = new DirectoryInfo(destinationFolderPath);
            var destinationFiles = destinationDirectory.GetFiles("*.xlsx")
                .Where(file => !file.Name.StartsWith("~"))
                .ToList();

            foreach (var destinationFile in destinationFiles)
            {
                using (var package = new ExcelPackage(destinationFile))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    int rowCount = worksheet.Dimension.Rows;
                    int destinationNameIndex = GetColumnIndexByColumnName(worksheet, destinationNames);
                    int destinationValueIndex = GetColumnIndexByColumnName(worksheet, destinationValues);

                    for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                    {
                        var nameCell = worksheet.Cells[rowIndex, destinationNameIndex];
                        string? name = nameCell.Value?.ToString();
                        string? shortName = Regex.Match(name ?? "", @"^(.*?)\s*\(")?.Groups[1]?.Value.Trim();

                        if (!string.IsNullOrEmpty(shortName))
                        {
                            string valuesString = GetSourceValues(sourceFolderPath, sourceNames, sourceValues, shortName);
                            var valueCell = worksheet.Cells[rowIndex, destinationValueIndex];
                            valueCell.Value = valuesString;
                        }
                    }

                    package.Save();
                }
            }
        }

        private string GetSourceValues(string sourceFolderPath, string sourceNames, string sourceValues, string shortName)
        {
            var sourceFiles = Directory.GetFiles(sourceFolderPath, "*.xlsx");

            foreach (var file in sourceFiles)
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    int rowCount = worksheet.Dimension.Rows;
                    int namesColumnIndex = GetColumnIndexByColumnName(worksheet, sourceNames);
                    int valuesColumnIndex = GetColumnIndexByColumnName(worksheet, sourceValues);

                    for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                    {
                        var nameCell = worksheet.Cells[rowIndex, namesColumnIndex];
                        var valueCell = worksheet.Cells[rowIndex, valuesColumnIndex];

                        string? name = nameCell.Value?.ToString();
                        string? value = valueCell.Value?.ToString();

                        if (name == "Gavia stellata")
                        {

                        }

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                        {
                            string normalizedShortName = Regex.Replace(shortName, @"[^a-zA-Z0-9]", "");
                            string normalizedComparisonString = Regex.Replace("Gavia stellata", @"[^a-zA-Z0-9]", "");

                            if (name.Equals(shortName, StringComparison.CurrentCultureIgnoreCase))
                            {
                                return value;
                            }
                        }
                    }
                }
            }

            return string.Empty;
        }

        private int GetColumnIndexByColumnName(ExcelWorksheet worksheet, string columnName)
        {
            var formattedName = columnName.ToLower().Trim();

            int columnCount = worksheet.Dimension.Columns;

            for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                var cellValue = worksheet.Cells[1, columnIndex].Value?.ToString()?.ToLower()?.Trim();
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Replace("\n", " ").Equals(formattedName, StringComparison.OrdinalIgnoreCase))
                {
                    return columnIndex;
                }
            }

            throw new ArgumentException($"Column '{columnName}' not found in the worksheet.");
        }
    }
}
