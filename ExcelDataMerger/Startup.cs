string sourceFiles = @"C:\Users\nc\Documents\NKB\source";
string dataFiles = @"C:\Users\nc\Documents\NKB\destination";


var sourceColumnNames = "Scientific Name";
var sourceColumnValues = "T";
var destinationColumnNames = "Species LATIN AND ENG";
var destinationColumnValues = "Type of presence";

var npoiManager = new NPOIManager();
var epplusManager = new EPPlusManager();
var xssfManager = new EPPlusManager();

xssfManager.GetNamesWithValues(sourceFiles, sourceColumnNames, sourceColumnValues);

Console.WriteLine("--------------------------------------------");

//epplusManager.GetNamesWithValues(dataFiles, destinationColumnNames, destinationColumnValues);
//npoiManager.GetNamesWithValues(dataFiles, destinationColumnNames, destinationColumnValues);
xssfManager.GetNamesWithValues(dataFiles, destinationColumnNames, destinationColumnValues);

Console.WriteLine("Excel file updated successfully.");