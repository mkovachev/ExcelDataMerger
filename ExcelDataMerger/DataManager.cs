
using OfficeOpenXml;

public class DataManager
{
    public void RetrieveScientificNamesAndValues(string sourceFolderPath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        DirectoryInfo sourceDirectory = new DirectoryInfo(sourceFolderPath);
        FileInfo[] sourceFiles = sourceDirectory.GetFiles("*.xlsx");

        Dictionary<string, List<string>> scientificNameValues = new Dictionary<string, List<string>>();

        foreach (FileInfo sourceFile in sourceFiles)
        {
            using (ExcelPackage sourcePackage = new ExcelPackage(sourceFile))
            {
                ExcelWorksheet sourceWorksheet = sourcePackage.Workbook.Worksheets.FirstOrDefault();

                if (sourceWorksheet != null)
                {
                    int rowCount = sourceWorksheet.Dimension?.Rows ?? 0;
                    int scientificNameColumnIndex = GetScientificNameColumnIndex(sourceWorksheet, "Scientific Name");
                    int valueColumnIndex = GetColumnIndexByName(sourceWorksheet, "T");

                    for (int rowIndex = 3; rowIndex <= rowCount; rowIndex++)
                    {
                        string scientificName = sourceWorksheet.Cells[rowIndex, scientificNameColumnIndex]?.Value?.ToString();
                        string value = sourceWorksheet.Cells[rowIndex, valueColumnIndex]?.Value?.ToString();

                        if (!string.IsNullOrEmpty(scientificName) && !string.IsNullOrEmpty(value))
                        {
                            if (!scientificNameValues.ContainsKey(scientificName))
                            {
                                scientificNameValues.Add(scientificName, new List<string>());
                            }

                            scientificNameValues[scientificName].Add(value);
                        }
                    }
                }
            }
        }

        foreach (var entry in scientificNameValues)
        {
            string name = entry.Key;
            List<string> values = entry.Value;

            string valuesString = string.Join(",", values);
            Console.WriteLine($"{name}: {valuesString}");
        }
    }

    private int GetScientificNameColumnIndex(ExcelWorksheet worksheet, string columnName)
    {
        int columnIndex = -1;

        int headerRow = 1; // Assuming the header row is the first row

        for (int column = 1; column <= worksheet.Dimension.Columns; column++)
        {
            var cellValue = worksheet.Cells[headerRow, column].Value?.ToString();

            if (cellValue != null && cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                columnIndex = column;
                break;
            }
        }

        return columnIndex;
    }

    private int GetColumnIndexByName(ExcelWorksheet worksheet, string columnName)
    {
        int columnIndex = -1;

        int headerRow = 1; // Assuming the header row is the first row

        for (int column = 1; column <= worksheet.Dimension.Columns; column++)
        {
            var cellValue = worksheet.Cells[headerRow, column].Value?.ToString();

            if (cellValue != null && cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                columnIndex = column;
                break;
            }
        }

        return columnIndex;
    }

}
