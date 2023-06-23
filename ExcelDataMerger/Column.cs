namespace ExcelDataMerger
{
    public class Column
    {
        public Column(string name, int index)
        {
            Name = name;
            Index = index;
        }

        public string? Name { get; set; }
        public int Index { get; set; }

    }
}
