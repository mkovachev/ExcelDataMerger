using OfficeOpenXml;

namespace ExcelDataMerger
{
    public class MergerManager
    {
        private readonly string folderPath;
        private readonly int animalTypeColumnIndex;

        public MergerManager(string folderPath, int animalTypeColumnIndex)
        {
            this.folderPath = folderPath;
            this.animalTypeColumnIndex = animalTypeColumnIndex;
        }

        public void MergeFiles(string outputFilePath)
        {
            var fileNames = Directory.GetFiles(folderPath, "*.xlsx");

            if (fileNames.Length == 0)
            {
                Console.WriteLine("No Excel files found in the specified folder.");
                return;
            }

            var animalTypes = new List<string>();
            var mergedData = new List<Animal>();

            foreach (string fileName in fileNames)
            {
                FileInfo fileInfo = new(fileName);

                using ExcelPackage package = new(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                int lastRow = worksheet.Dimension.End.Row;

                for (int row = 2; row <= lastRow; row++)
                {
                    string? animalType = worksheet.Cells[row, animalTypeColumnIndex]?.Value?.ToString();

                    if (!string.IsNullOrEmpty(animalType))
                    {
                        if (!animalTypes.Contains(animalType))
                        {
                            animalTypes.Add(animalType);
                        }

                        var animal = new Animal()
                        {
                            Type = animalType,
                            Name = worksheet.Cells[row, 1]?.Value?.ToString(),
                            Age = Convert.ToInt32(worksheet.Cells[row, 2]?.Value)
                        };

                        mergedData.Add(animal);
                    }
                }
            }

            using ExcelPackage mergedPackage = new();
            ExcelWorksheet mergedWorksheet = mergedPackage.Workbook.Worksheets.Add("Merged Data");

            mergedWorksheet.Cells[1, 1].Value = "Animal Type";
            mergedWorksheet.Cells[1, 2].Value = "Name";
            mergedWorksheet.Cells[1, 3].Value = "Age";

            var currentRow = 2;

            foreach (Animal animalData in mergedData)
            {
                mergedWorksheet.Cells[currentRow, 1].Value = animalData.Type;
                mergedWorksheet.Cells[currentRow, 2].Value = animalData.Name;
                mergedWorksheet.Cells[currentRow, 3].Value = animalData.Age;

                currentRow++;
            }

            mergedPackage.SaveAs(new FileInfo(outputFilePath));
        }
    }
}