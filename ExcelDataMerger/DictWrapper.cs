namespace ExcelDataMerger
{
    public class DictWrapper : Dictionary<string, List<string>>
    {
        private readonly IEqualityComparer<string> _comparer;

        public DictWrapper(IEqualityComparer<string> comparer)
        {
            _comparer = comparer ?? throw new ArgumentNullException(nameof(comparer));
        }

        public new bool ContainsKey(string key)
        {
            return Keys.Any(k => _comparer.Equals(k, key));
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                foreach (var pair in this)
                {
                    hash = ((hash * 23) + (_comparer.GetHashCode(pair.Key) * 397)) ^ (pair.Value != null ? pair.Value.GetHashCode() : 0);
                }
                return hash;
            }
        }
    }
}
