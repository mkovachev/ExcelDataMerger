using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class XSSFManager
{
    private Dictionary<string, List<string>> GetSourceData(string sourceFolderPath, string sourceNames, string sourceValues)
    {
        var sourceFiles = Directory.GetFiles(sourceFolderPath, "*.xlsx");

        Dictionary<string, List<string>> names = new Dictionary<string, List<string>>();

        foreach (var file in sourceFiles)
        {
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                var workbook = new XSSFWorkbook(stream);
                var sheet = workbook.GetSheetAt(0); // Assuming data is in the first sheet

                int rowCount = sheet.LastRowNum + 1;
                int namesColumnIndex = GetIndexByColumnName(sheet, sourceNames);
                int valuesColumnIndex = GetIndexByColumnName(sheet, sourceValues);

                for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        string? name = row.GetCell(namesColumnIndex)?.ToString();
                        string? value = row.GetCell(valuesColumnIndex)?.ToString();

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                        {
                            if (!names.ContainsKey(name))
                            {
                                names.Add(name, new List<string>());
                            }

                            names[name].Add(value);
                        }
                    }
                }
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

    public void UpdateTypeOfPresence(string sourceFolderPath, string destinationFolderPath, string sourceNames, string sourceValues, string destinationNames, string destinationValues)
    {
        var sourceData = GetSourceData(sourceFolderPath, sourceNames, sourceValues);

        var destinationDirectory = new DirectoryInfo(destinationFolderPath);
        var destinationFiles = destinationDirectory.GetFiles("*.xlsx");

        foreach (var destinationFile in destinationFiles)
        {
            if (destinationFile.Name.StartsWith("~$"))
                continue;

            using (var stream = new FileStream(destinationFile.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                var workbook = new XSSFWorkbook(stream);
                var sheet = workbook.GetSheetAt(0); // Assuming data is in the first sheet

                int rowCount = sheet.LastRowNum + 1;
                int destinationNameIndex = GetIndexByColumnName(sheet, destinationNames);
                int destinationValueIndex = GetIndexByColumnName(sheet, destinationValues);

                for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        string? name = row.GetCell(destinationNameIndex)?.ToString();

                        string shortName = string.Empty;
                        if (name is not null)
                        {
                            string[]? nameParts = name?.Split(' ');
                            if (nameParts.Length > 1)
                            {
                                shortName = $"{nameParts?[0]} {nameParts?[1]}".Trim();
                            }
                            else
                            {
                                shortName = nameParts[0].Trim();
                            }
                        }
                        else
                        {
                            shortName = string.Empty;
                        }

                        // TODO: 
                        if (!string.IsNullOrEmpty(shortName) && sourceData.Keys.Any(key => key.Equals(shortName, StringComparison.OrdinalIgnoreCase)))
                        {
                            var values = sourceData[shortName];
                            string valuesString = string.Join(",", values);

                            row.GetCell(destinationNameIndex)?.SetCellValue(valuesString);
                        }
                    }
                }

                using (var writeStream = new FileStream(destinationFile.FullName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(writeStream);
                }
            }
        }
    }
}
