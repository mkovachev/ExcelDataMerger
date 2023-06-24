using ExcelDataMerger;

string sourceFolderPath = @"C:\Users\nc\Documents\NKB\source";
string destinationFolderPath = @"C:\Users\nc\Documents\NKB\destination";

var logManager = new LogManager(destinationFolderPath);
logManager.ClearLog();

// case Type of presence
string sourceNames = "Scientific Name";
string sourceValues = "T";
string destinationSpecies = "Species LATIN AND ENG";
string destinationTypeOfPresence = "Type of presence";

// case Site assessment
string destinationSiteAssessmentValues = "Site assessment: con / glo";
string sourceCon = "Con.";
string sourceGlo = "Glo.";

// Check and add columns process
var columnManager = new ColumnManager(logManager);

var columns = new List<Column>()
{
    new Column { Name = destinationTypeOfPresence, Index = 6 },
    new Column { Name = destinationSiteAssessmentValues, Index = 11 }
};

int totalFilesProcessed = 0;
int filesWithAddedColumns = 0;
int filesWithExistingColumns = 0;

var destinationDirectory = new DirectoryInfo(destinationFolderPath);
var destinationFiles = destinationDirectory.GetFiles("*.xlsx");

foreach (var destinationFile in destinationFiles)
{
    if (destinationFile.Name.StartsWith("~$"))
        continue;

    columnManager.LoadExcelFile(destinationFile);
    columnManager.CreateColumns(columns, destinationSpecies, destinationTypeOfPresence);

    var fileName = destinationFile.Name;
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

// Update Columns process
var npoiManager = new NPOIManager(logManager);
logManager.Log("----------------------------------------------------------------");
npoiManager.UpdateColumns(sourceFolderPath, destinationFolderPath, sourceNames, sourceValues, destinationSpecies, destinationTypeOfPresence);
logManager.Log("----------------------------------------------------------------");
npoiManager.UpdateSiteAssessmentColumn(sourceFolderPath, destinationFolderPath, sourceNames, sourceCon, sourceGlo, destinationSpecies, destinationSiteAssessmentValues);
