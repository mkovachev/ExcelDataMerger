using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class XSSFManager
{
    public void GetNamesWithValues(string sourceFolderPath, string columnNames, string columnsValues)
    {
        DirectoryInfo directory = new DirectoryInfo(sourceFolderPath);
        FileInfo[] files = directory.GetFiles("*.xlsx");

        Dictionary<string, List<string>> names = new Dictionary<string, List<string>>();

        foreach (FileInfo file in files)
        {
            using (FileStream stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0); // Assuming data is in the first sheet

                int rowCount = sheet.LastRowNum + 1;
                int namesColumnIndex = GetIndexByColumnName(sheet, columnNames);
                int valuesColumnIndex = GetIndexByColumnName(sheet, columnsValues);

                for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
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

        foreach (var entry in names)
        {
            string name = entry.Key;
            List<string> values = entry.Value;

            string valuesString = string.Join(",", values);
            Console.WriteLine($"{name}: {valuesString}");
        }
    }

    private int GetIndexByColumnName(ISheet sheet, string columnName)
    {
        var formattedColumnName = columnName.ToLower().Trim();

        int columnIndex = -1;
        IRow headerRow = sheet.GetRow(0);

        for (int column = 0; column < headerRow.LastCellNum; column++)
        {
            string? cellValue = headerRow.GetCell(column)?.ToString()?.ToLower()?.Trim();

            if (!string.IsNullOrEmpty(cellValue) && cellValue.Replace("\n", " ").Equals(formattedColumnName, StringComparison.OrdinalIgnoreCase))
            {
                columnIndex = column;
                break;
            }
        }

        return columnIndex;
    }
}
