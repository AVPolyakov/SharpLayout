namespace Examples
{
    public class Indexer
    {
        private readonly string _value;
        private int _index;

        public Indexer(string value) => _value = value;

        public void Next() => _index++;

        public string Current => _index < _value?.Length ? _value.Substring(_index, 1) : "";
        
        public static implicit operator Indexer(string value) => new Indexer(value);
    }
}