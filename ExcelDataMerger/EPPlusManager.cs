
using OfficeOpenXml;

public class EPPlusManager
{
    public void GetNamesWithValues(string sourceFolderPath, string columnNames, string columnsValues)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        DirectoryInfo directory = new DirectoryInfo(sourceFolderPath);
        FileInfo[] files = directory.EnumerateFiles("*.xlsx")
            .Where(file => !file.Name.StartsWith("~$"))
            .ToArray();

        Dictionary<string, List<string>> names = new Dictionary<string, List<string>>();

        foreach (FileInfo file in files)
        {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int namesColumnIndex = GetIndexByColumnName(worksheet, columnNames);
                    int valuesColumnIndex = GetIndexByColumnName(worksheet, columnsValues);

                    for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                    {
                        string name = worksheet.Cells[rowIndex, namesColumnIndex]?.Value?.ToString();
                        string value = worksheet.Cells[rowIndex, valuesColumnIndex]?.Value?.ToString();

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

        foreach (var entry in names)
        {
            string name = entry.Key;
            List<string> values = entry.Value;

            string valuesString = string.Join(",", values);
            Console.WriteLine($"{name}: {valuesString}");
        }
    }

    private int GetIndexByColumnName(ExcelWorksheet worksheet, string columnName)
    {
        var formattedColumnName = columnName.ToLower().Trim();
        int columnIndex = -1;

        int headerRow = 1; // Assuming the header row is the first row

        for (int column = 1; column <= worksheet.Dimension.Columns; column++)
        {
            var cellValue = worksheet.Cells[headerRow, column].Value?.ToString().ToLower().Trim();

            if (cellValue != null && cellValue.Replace("\n", " ").Equals(formattedColumnName, StringComparison.OrdinalIgnoreCase))
            {
                columnIndex = column;
                break;
            }
        }

        return columnIndex;
    }
}
