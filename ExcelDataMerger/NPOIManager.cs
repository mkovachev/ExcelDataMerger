﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

public class NPOIManager
{
    private readonly LogManager logManager;

    public NPOIManager(LogManager logManager)
    {
        this.logManager = logManager;
    }
    public void UpdateSiteAssessmentColumn(string sourceFolderPath, string destinationFolderPath, string sourceNames, string sourceCon, string sourceGlo, string destinationSpecies, string destinationSiteAssessmentValues)
    {
        try
        {
            var sourceData = GetSecondSourceData(sourceFolderPath, sourceNames, sourceCon, sourceGlo);

            var destinationDirectory = new DirectoryInfo(destinationFolderPath);
            var destinationFiles = destinationDirectory.GetFiles("*.xlsx");

            int totalUpdatedValues = 0;
            int totalUpdatedFiles = 0;

            foreach (var destinationFile in destinationFiles)
            {
                try
                {
                    if (destinationFile.Name.StartsWith("~$"))
                        continue;

                    int updatedValuesCount = 0;

                    using (var stream = new FileStream(destinationFile.FullName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        var workbook = new XSSFWorkbook(stream);
                        var sheet = workbook.GetSheetAt(0); // Assuming data is in the first sheet

                        int rowCount = sheet.LastRowNum + 1;
                        int namesColumnIndex = GetIndexByFirstColumnName(sheet, destinationSpecies);
                        int siteAssessmentColumnIndex = GetIndexByFirstColumnName(sheet, destinationSiteAssessmentValues);

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
                                        string conValuesString = string.Join(",", values["con"].Distinct());
                                        string gloValuesString = string.Join(",", values["glo"].Distinct());

                                        var siteAssessmentCell = row.GetCell(siteAssessmentColumnIndex);
                                        if (siteAssessmentCell == null)
                                            siteAssessmentCell = row.CreateCell(siteAssessmentColumnIndex, CellType.String);
                                        siteAssessmentCell.SetCellValue($"{conValuesString} / {gloValuesString}");

                                        logManager.Log($"{name}: {conValuesString} / {gloValuesString}");
                                        updatedValuesCount++;
                                        totalUpdatedValues++;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                logManager.Log($"Error updating row {rowIndex + 1} in file '{destinationFile.Name}': {ex.Message}");
                            }
                        }

                        if (updatedValuesCount > 0)
                        {
                            totalUpdatedFiles++;
                            logManager.Log($"{destinationFile.Name}: {updatedValuesCount}");

                        }

                        using (var writeStream = new FileStream(destinationFile.FullName, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(writeStream);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logManager.Log($"Error processing file '{destinationFile.Name}': {ex.Message}");
                }
            }

            logManager.Log($"Total files updated: {totalUpdatedFiles}");
            logManager.Log("----------------------------------------------------------------");
        }
        catch (Exception ex)
        {
            logManager.Log($"An error occurred during the site assessment column update process: {ex.Message}");
        }
    }



    private Dictionary<string, Dictionary<string, List<string>>> GetSecondSourceData(string sourceFolderPath, string sourceNames, string sourceCon, string sourceGlo)
    {
        var sourceDirectory = new DirectoryInfo(sourceFolderPath);
        var sourceFiles = sourceDirectory.GetFiles("*.xlsx");

        var names = new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.CurrentCultureIgnoreCase);

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
                    int namesColumnIndex = GetIndexByFirstColumnName(sheet, sourceNames);
                    int conColumnIndex = GetIndexBySecondColumnName(sheet, sourceCon);
                    int gloColumnIndex = GetIndexBySecondColumnName(sheet, sourceGlo);

                    for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                    {
                        try
                        {
                            var row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                string name = GetCellValue(row.GetCell(namesColumnIndex));
                                string conValue = GetCellValue(row.GetCell(conColumnIndex));
                                string gloValue = GetCellValue(row.GetCell(gloColumnIndex));

                                if (!string.IsNullOrEmpty(name) && (!string.IsNullOrEmpty(conValue) || !string.IsNullOrEmpty(gloValue)))
                                {
                                    if (!names.ContainsKey(name))
                                        names.Add(name, new Dictionary<string, List<string>>());

                                    if (!names[name].ContainsKey("con"))
                                        names[name].Add("con", new List<string>());

                                    if (!names[name].ContainsKey("glo"))
                                        names[name].Add("glo", new List<string>());

                                    if (!string.IsNullOrEmpty(conValue))
                                        names[name]["con"].Add(conValue);

                                    if (!string.IsNullOrEmpty(gloValue))
                                        names[name]["glo"].Add(gloValue);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logManager.Log($"Error reading row {rowIndex + 1} in file '{file.Name}': {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logManager.Log($"Error processing file '{file.Name}': {ex.Message}");
            }
        }

        return names;
    }


    private int GetIndexBySecondColumnName(ISheet sheet, string columnName)
    {
        var formattedName = columnName.ToLower().Trim();

        int columnIndex = -1;
        var headerRow = sheet.GetRow(2);

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
                        int namesColumnIndex = GetIndexByFirstColumnName(sheet, destinationNames);
                        int valueColumnIndex = GetIndexByFirstColumnName(sheet, destinationValues);

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
                                        string valuesString = string.Join(",", values.Distinct());

                                        var valueCell = row.GetCell(valueColumnIndex);
                                        if (valueCell == null)
                                            valueCell = row.CreateCell(valueColumnIndex, CellType.String);
                                        valueCell.SetCellValue(valuesString);

                                        logManager.Log($"{name}: {valuesString}");
                                        updatedNamesCount++;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                logManager.Log($"Error updating row {rowIndex + 1} in file '{destinationFile.Name}': {ex.Message}");
                            }
                        }

                        if (updatedNamesCount > 0)
                        {
                            totalUpdatedNames += updatedNamesCount;
                            totalUpdatedFiles++;
                            logManager.Log($"{destinationFile.Name}: {updatedNamesCount}");
                        }

                        using (var writeStream = new FileStream(destinationFile.FullName, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(writeStream);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logManager.Log($"Error processing file '{destinationFile.Name}': {ex.Message}");
                }
            }

            logManager.Log("Type of presence Summary:");
            logManager.Log($"Total names updated per file: {totalUpdatedNames}");
            logManager.Log($"Total files updated: {totalUpdatedFiles}");
            logManager.Log("----------------------------------------------------------------");
        }
        catch (Exception ex)
        {
            logManager.Log($"An error occurred during the column update process: {ex.Message}");
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
                    int namesColumnIndex = GetIndexByFirstColumnName(sheet, sourceNames);
                    int valuesColumnIndex = GetIndexByFirstColumnName(sheet, sourceValues);

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
                            logManager.Log($"Error reading row {rowIndex + 1} in file '{file.Name}': {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logManager.Log($"Error processing file '{file.Name}': {ex.Message}");
            }
        }

        return names;
    }

    private int GetIndexByFirstColumnName(ISheet sheet, string columnName)
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
