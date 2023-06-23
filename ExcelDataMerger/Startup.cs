using ExcelDataMerger;

string directoryPath = @"C:\Users\nc\Documents\NKB\data";
string[] files = Directory.GetFiles(directoryPath, "*.xlsx");

List<Column> columns = new List<Column>();
columns.Add(new Column("Type of presence", 6));
columns.Add(new Column("Site assessment: con / glo", 11));

var columnManager = new ColumnManager();
int totalFilesProcessed = 0;
int filesWithAddedColumns = 0;
int filesWithExistingColumns = 0;

foreach (string filePath in files)
{
    columnManager.LoadExcelFile(filePath);
    columnManager.CreateColumns(columns);

    string fileName = Path.GetFileName(filePath);
    columnManager.PrintFileSummary(fileName);

    columnManager.Reset();

    totalFilesProcessed++;
    if (columnManager.ColumnsAdded > 0)
    {
        filesWithAddedColumns++;
    }
    if (columnManager.ColumnsExist > 0)
    {
        filesWithExistingColumns++;
    }
}

columnManager.PrintSummary(totalFilesProcessed, filesWithAddedColumns, filesWithExistingColumns);
