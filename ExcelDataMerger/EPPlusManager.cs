﻿using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace ExcelDataMerger
{
    public class EPPlusManager
    {
        public void UpdateColumns(string sourceFolderPath, string destinationFolderPath, string sourceNames, string sourceValues, string destinationNames, string destinationValues)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var sourceData = GetSourceData(sourceFolderPath, sourceNames, sourceValues);
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

                        if (!string.IsNullOrEmpty(shortName) && sourceData.ContainsKey(shortName))
                        {
                            var values = sourceData[shortName];
                            string valuesString = string.Join(",", values);

                            var valueCell = worksheet.Cells[rowIndex, destinationValueIndex];
                            valueCell.Value = valuesString;
                        }
                    }

                    package.Save();
                }
            }
        }

        private Dictionary<string, List<string>> GetSourceData(string sourceFolderPath, string sourceNames, string sourceValues)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var sourceFiles = Directory.GetFiles(sourceFolderPath, "*.xlsx");

            var sourceData = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

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

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                        {
                            if (!sourceData.ContainsKey(name))
                            {
                                sourceData[name] = new List<string>();
                            }

                            sourceData[name].Add(value);
                        }
                    }
                }
            }

            return sourceData;
        }

        private int GetColumnIndexByColumnName(ExcelWorksheet worksheet, string columnName)
        {
            var formattedName = columnName.ToLower().Trim();

            int columnCount = worksheet.Dimension.Columns;

            for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                var cellValue = worksheet.Cells[1, columnIndex].Value?.ToString()?.ToLower()?.Trim(); ;
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Replace("\n", " ").Equals(formattedName, StringComparison.OrdinalIgnoreCase))
                {
                    return columnIndex;
                }
            }

            throw new ArgumentException($"Column '{columnName}' not found in the worksheet.");
        }
    }
}
