using ExcelDataMerger;

public class Program
{
    public static void Main()
    {
        string sourceFolderPath = @"C:\Users\nc\Documents\NKB\source";
        string destinationFolderPath = @"C:\Users\nc\Documents\NKB\destination";

        // case Type of presence
        string sourceNames = "Scientific Name";
        string sourceValues = "T";
        string destinationSpecies = "Species LATIN AND ENG";
        string destinationTypeOfPresense = "Type of presence";

        // case Site assessment
        string destinationSiteAssessmentValues = "Site assessment: con / glo";

        // Check and add columns process
        var columnManager = new ColumnManager();

        var columns = new List<Column>()
        {
            new Column { Name = destinationTypeOfPresense, Index = 6 },
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
            columnManager.CreateColumns(columns, destinationSpecies, destinationTypeOfPresense);

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
        var npoiManager = new NPOIManager();
        npoiManager.UpdateColumns(sourceFolderPath, destinationFolderPath, sourceNames, sourceValues, destinationSpecies, destinationTypeOfPresense);

        Console.WriteLine("Updating finished.");
    }
}
