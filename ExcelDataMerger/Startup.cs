string sourceFiles = @"C:\Users\nc\Documents\NKB\source";

DataManager dataManager = new DataManager();
dataManager.RetrieveScientificNamesAndValues(sourceFiles);
//dataManager.UpdateExcelFileWithPresenceType(sourceFiles, destinationFiles);

Console.WriteLine("Excel file updated successfully.");