using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class NPOIManager
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
                ISheet sheet = workbook.GetSheetAt(0);

                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum + 1;
                    int namesColumnIndex = GetIndexByColumnName(sheet.GetRow(0), columnNames);
                    int valuesColumnIndex = GetIndexByColumnName(sheet.GetRow(0), columnsValues);

                    for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);

                        if (row != null)
                        {
                            string name = GetCellValue(row.GetCell(namesColumnIndex));
                            string value = GetCellValue(row.GetCell(valuesColumnIndex));

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
        }

        foreach (var entry in names)
        {
            string name = entry.Key;
            List<string> values = entry.Value;

            string valuesString = string.Join(",", values);
            Console.WriteLine($"{name}: {valuesString}");
        }
    }

    private int GetIndexByColumnName(IRow headerRow, string columnName)
    {
        var formattedColumnName = columnName.ToLower().Trim();

        int columnIndex = -1;

        if (headerRow != null)
        {
            for (int column = 0; column < headerRow.LastCellNum; column++)
            {
                string cellValue = GetCellValue(headerRow.GetCell(column))?.ToLower()?.Trim();

                if (cellValue != null && cellValue.Replace("\n", " ").Equals(formattedColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = column;
                    break;
                }
            }
        }

        return columnIndex;
    }

    private string? GetCellValue(ICell cell)
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
                    // Handle formula cells if needed
                    return cell.CellFormula;
                default:
                    return null;
            }
        }

        return null;
    }
}
