﻿using ExcelDataMerger;
using OfficeOpenXml;

public class ColumnManager
{
    private ExcelPackage? package;
    private ExcelWorksheet? worksheet;
    private int columnsAdded;
    private int columnsExist;
    private List<string>? addedColumns;
    private List<string>? existingColumns;

    public void LoadExcelFile(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        package = new ExcelPackage(new FileInfo(filePath));
        worksheet = package.Workbook.Worksheets[0];
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

            worksheet.Cells[1, column.Index].Value = column.Name;
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

            if (!string.IsNullOrEmpty(cellValue) && cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
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
        Console.WriteLine("Summary:");
        Console.WriteLine($"Total files processed: {totalFilesProcessed}");
        Console.WriteLine($"Total files with added columns: {filesWithAddedColumns}");
        Console.WriteLine($"Total files with existing columns: {filesWithExistingColumns}");
    }

    public void PrintFileSummary(string processedFile)
    {
        Console.WriteLine($"Processed file: {processedFile}");

        if (addedColumns.Count > 0)
        {
            Console.WriteLine("Columns added:");
            foreach (string columnName in addedColumns)
            {
                Console.WriteLine(columnName);
            }
        }

        if (existingColumns.Count > 0)
        {
            Console.WriteLine("Columns already present:");
            foreach (string columnName in existingColumns)
            {
                Console.WriteLine(columnName);
            }
        }

        Console.WriteLine();
    }

    public int ColumnsAdded => columnsAdded;
    public int ColumnsExist => columnsExist;
}