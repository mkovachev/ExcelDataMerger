using ExcelDataMerger;

string folderPath = @"C:\YourFolderPath\";
int animalTypeColumnIndex = 3;
string outputFilePath = @"C:\YourOutputFilePath\output.xlsx";

MergerManager mergerManager = new(folderPath, animalTypeColumnIndex);
mergerManager.MergeFiles(outputFilePath);

Console.WriteLine("Data merged successfully!");
Console.ReadLine();