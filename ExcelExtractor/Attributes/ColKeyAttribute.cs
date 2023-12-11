namespace ExcelExtractor.Attributes
{
    public class ColKeyAttribute : Attribute
    {
        private readonly string _key;
        public string Key => _key.ToLower();
        public ColKeyAttribute(string key) => _key = key;
    }
}
