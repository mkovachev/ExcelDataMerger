using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

public class NPOIManager
{
    public void UpdateColumns(string sourceFolderPath, string destinationFolderPath, string sourceNames, string sourceValues, string destinationNames, string destinationValues)
    {
        try
        {
            var sourceData = GetSourceData(sourceFolderPath, sourceNames, sourceValues);

            var destinationDirectory = new DirectoryInfo(destinationFolderPath);
            var destinationFiles = destinationDirectory.GetFiles("*.xlsx");

            int totalUpdatedNames = 0;
            int totalUpdatedFiles = 0;

            foreach (var destinationFile in destinationFiles)
            {
                try
                {
                    if (destinationFile.Name.StartsWith("~$"))
                        continue;

                    using (var stream = new FileStream(destinationFile.FullName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        var workbook = new XSSFWorkbook(stream);
                        var sheet = workbook.GetSheetAt(0); // Assuming data is in the first sheet

                        int rowCount = sheet.LastRowNum + 1;
                        int namesColumnIndex = GetIndexByColumnName(sheet, destinationNames);
                        int valueColumnIndex = GetIndexByColumnName(sheet, destinationValues);

                        int updatedNamesCount = 0;

                        for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                        {
                            try
                            {
                                var row = sheet.GetRow(rowIndex);
                                if (row != null)
                                {
                                    string name = GetCellValue(row.GetCell(namesColumnIndex));
                                    string? shortName = Regex.Match(name, @"^(.*?)\s*\(")?.Groups[1]?.Value.Trim();

                                    if (!string.IsNullOrEmpty(shortName) && sourceData.ContainsKey(shortName))
                                    {
                                        var values = sourceData[shortName];
                                        string valuesString = string.Join(",", values);

                                        var valueCell = row.GetCell(valueColumnIndex);
                                        if (valueCell == null)
                                            valueCell = row.CreateCell(valueColumnIndex, CellType.String);
                                        valueCell.SetCellValue(valuesString);

                                        Console.WriteLine($"Updated value for name '{name}': {valuesString}");
                                        updatedNamesCount++;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error updating row {rowIndex + 1} in file '{destinationFile.Name}': {ex.Message}");
                            }
                        }

                        if (updatedNamesCount > 0)
                        {
                            totalUpdatedNames += updatedNamesCount;
                            totalUpdatedFiles++;
                            Console.WriteLine($"Total updated names in file '{destinationFile.Name}': {updatedNamesCount}");
                            Console.WriteLine();
                        }

                        using (var writeStream = new FileStream(destinationFile.FullName, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(writeStream);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{destinationFile.Name}': {ex.Message}");
                }
            }

            Console.WriteLine("Update Columns Summary:");
            Console.WriteLine($"Total updated names across all files: {totalUpdatedNames}");
            Console.WriteLine($"Total files with updated names: {totalUpdatedFiles}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during the column update process: {ex.Message}");
        }
    }

    private Dictionary<string, List<string>> GetSourceData(string sourceFolderPath, string sourceNames, string sourceValues)
    {
        var sourceDirectory = new DirectoryInfo(sourceFolderPath);
        var sourceFiles = sourceDirectory.GetFiles("*.xlsx");

        var names = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);

        foreach (var file in sourceFiles)
        {
            try
            {
                if (file.Name.StartsWith("~$"))
                    continue;

                using (var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new XSSFWorkbook(stream);
                    var sheet = workbook.GetSheetAt(0); // Assuming data is in the first sheet

                    int rowCount = sheet.LastRowNum + 1;
                    int namesColumnIndex = GetIndexByColumnName(sheet, sourceNames);
                    int valuesColumnIndex = GetIndexByColumnName(sheet, sourceValues);

                    for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                    {
                        try
                        {
                            var row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                string name = GetCellValue(row.GetCell(namesColumnIndex));
                                string value = GetCellValue(row.GetCell(valuesColumnIndex));

                                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                                {
                                    if (!names.ContainsKey(name))
                                        names.Add(name, new List<string>());

                                    names[name].Add(value);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error reading row {rowIndex + 1} in file '{file.Name}': {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{file.Name}': {ex.Message}");
            }
        }

        return names;
    }


    private int GetIndexByColumnName(ISheet sheet, string columnName)
    {
        var formattedName = columnName.ToLower().Trim();

        int columnIndex = -1;
        var headerRow = sheet.GetRow(0);

        for (int column = 0; column < headerRow.LastCellNum; column++)
        {
            var cellValue = headerRow.GetCell(column)?.ToString()?.ToLower()?.Trim();

            if (!string.IsNullOrEmpty(cellValue) && cellValue.Replace("\n", " ").Equals(formattedName, StringComparison.OrdinalIgnoreCase))
            {
                columnIndex = column;
                break;
            }
        }

        return columnIndex;
    }

    private string GetCellValue(ICell cell)
    {
        if (cell != null)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    return cell.CellFormula;
                default:
                    return string.Empty;
            }
        }

        return string.Empty;
    }
}
