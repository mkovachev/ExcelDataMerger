using ExcelDataMerger;
using OfficeOpenXml;

public class ColumnManager
{
    private ExcelPackage? package;
    private ExcelWorksheet? worksheet;
    private int columnsAdded;
    private int columnsExist;
    private List<string>? addedColumns;
    private List<string>? existingColumns;

    private readonly LogManager logManager;

    public ColumnManager(LogManager logManager)
    {
        this.logManager = logManager;
    }

    public void LoadExcelFile(FileInfo filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        package = new ExcelPackage(filePath);
        worksheet = package.Workbook.Worksheets.FirstOrDefault();
    }

    public void CreateColumns(List<Column> columns)
    {
        addedColumns = new List<string>();
        existingColumns = new List<string>();

        foreach (Column column in columns)
        {
            if (CreateColumn(column))
            {
                addedColumns.Add(column.Name);
                columnsAdded++;
            }
            else
            {
                existingColumns.Add(column.Name);
                columnsExist++;
            }
        }

        package?.Save();
    }

    public void MergeCellsByNumber(List<Column> columns, int numberOfCells)
    {
        if (worksheet != null)
        {
            foreach (Column column in columns)
            {
                // Find the column index by name
                var cell = worksheet.Cells["1:1"].FirstOrDefault(c => c.Value?.ToString() == column.Name);
                if (cell != null)
                {
                    int columnIndex = cell.Start.Column;

                    // Merge the cells in the specified range
                    int startRow = 2; // Assuming data starts from the second row
                    int endRow = startRow + numberOfCells - 1;

                    ExcelRange range = worksheet.Cells[startRow, columnIndex, endRow, columnIndex];
                    range.Merge = true;
                }
                else
                {
                    Console.WriteLine($"Column '{column.Name}' not found in the worksheet.");
                }
            }

            // Save the changes to the Excel file
            package?.Save();
        }
        else
        {
            Console.WriteLine("Excel file is not loaded. Please call LoadExcelFile method before merging cells.");
        }
    }

    public void MergeCells(List<Column> columns, string destinationSpecies)
    {
        int speciesColumnIndex = GetColumnIndexByName(destinationSpecies);
        int speciesCellCount = worksheet.Dimension.Rows - 1;

        List<int> mergeStartIndexes = new List<int>();
        List<int> mergeEndIndexes = new List<int>();

        // Find consecutive cells with the same value in the destinationSpecies column
        int startRowIndex = 2;
        while (startRowIndex <= speciesCellCount + 1)
        {
            string currentSpecies = worksheet.Cells[startRowIndex, speciesColumnIndex].Text;
            int cellLength = 1;
            int nextRowIndex = startRowIndex + 1;

            // Check if the next cells have the same value
            while (nextRowIndex <= speciesCellCount + 1 && worksheet.Cells[nextRowIndex, speciesColumnIndex].Text == currentSpecies)
            {
                cellLength++;
                nextRowIndex++;
            }

            if (cellLength > 1)
            {
                mergeStartIndexes.Add(startRowIndex);
                mergeEndIndexes.Add(startRowIndex + cellLength - 1);
            }

            startRowIndex = nextRowIndex;
        }

        foreach (Column column in columns)
        {
            if (column.Name != destinationSpecies)
            {
                for (int i = 0; i < mergeStartIndexes.Count; i++)
                {
                    int startIndex = mergeStartIndexes[i];
                    int endRowIndex = mergeEndIndexes[i];

                    string startCellAddress = GetCellAddress(column.Index, startIndex);
                    string endCellAddress = GetCellAddress(column.Index, endRowIndex);

                    worksheet.Cells[startCellAddress + ":" + endCellAddress].Merge = true;
                }
            }
        }

        package?.Save();
    }

    private string GetCellAddress(int columnIndex, int rowIndex)
    {
        string columnName = GetColumnNameFromIndex(columnIndex);
        return columnName + rowIndex.ToString();
    }

    private string GetColumnNameFromIndex(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }



    private bool CreateColumn(Column column)
    {
        int existingColumnIndex = GetColumnIndexByName(column.Name ?? string.Empty);

        if (existingColumnIndex == -1)
        {
            int lastColumnIndex = worksheet?.Dimension?.Columns ?? 0;

            if (column.Index <= lastColumnIndex)
            {
                // Shift existing columns to the right
                worksheet?.InsertColumn(column.Index, 1);
            }

            // Copy formatting from the adjacent cell in the header row
            int adjacentColumnIndex = column.Index - 1;
            var adjacentCell = worksheet.Cells[1, adjacentColumnIndex];
            worksheet.Cells[1, column.Index].Value = column.Name;
            worksheet.Cells[1, column.Index].StyleID = adjacentCell.StyleID;

            return true;
        }
        else
        {
            return false;
        }
    }

    private int GetColumnIndexByName(string columnName)
    {
        int columnCount = worksheet?.Dimension?.Columns ?? 0;

        for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
        {
            string? cellValue = worksheet?.Cells[1, columnIndex]?.Value?.ToString();

            if (!string.IsNullOrEmpty(cellValue) && cellValue.Replace("\n", " ").Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return columnIndex;
            }
        }

        return -1;
    }

    public void Reset()
    {
        package?.Dispose();
        worksheet = null;
        columnsAdded = 0;
        columnsExist = 0;
    }

    public void PrintSummary(int totalFilesProcessed, int filesWithAddedColumns, int filesWithExistingColumns)
    {
        logManager.Log("Create column summary:");
        logManager.Log($"Total files processed: {totalFilesProcessed}");
        logManager.Log($"Total files with added columns: {filesWithAddedColumns}");
        logManager.Log($"Total files with existing columns: {filesWithExistingColumns}");
        logManager.Log("----------------------------------------------------------------");
    }

    public void PrintFileSummary(string processedFile)
    {
        logManager.Log($"Processed file: {processedFile}");

        if (addedColumns.Count > 0)
        {
            logManager.Log("Columns added:");
            foreach (string columnName in addedColumns)
            {
                logManager.Log(columnName);
            }
        }

        if (existingColumns.Count > 0)
        {
            logManager.Log("Columns already present:");
            foreach (string columnName in existingColumns)
            {
                logManager.Log(columnName);
            }
        }

        logManager.Log("----------------------------------------------------------------");
    }

    public int ColumnsAdded => columnsAdded;
    public int ColumnsExist => columnsExist;
}
