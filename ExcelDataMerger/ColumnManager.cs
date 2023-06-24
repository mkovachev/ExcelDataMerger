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

    public void CreateColumns(List<Column> columns, string destinationNames, string destinationColumn)
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

        ///MergeCellsInNewColumns(destinationNames, destinationColumn);

        package?.Save();
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
