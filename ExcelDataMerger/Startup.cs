public class Program
{
    public static void Main()
    {
        string sourceFolderPath = @"C:\Users\nc\Documents\NKB\source";
        string destinationFilePath = @"C:\Users\nc\Documents\NKB\destination";

        string sourceNames = "Scientific Name";
        string sourceValues = "T";
        string destinationNames = "Species LATIN AND ENG";
        string destinationValues = "Type of presence";

        var xssfManager = new XSSFManager();
        xssfManager.UpdateTypeOfPresence(sourceFolderPath, destinationFilePath, sourceNames, sourceValues, destinationNames, destinationValues);

        Console.WriteLine("Excel file updated successfully.");
    }
}
