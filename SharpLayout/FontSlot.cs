using System.Collections.Concurrent;
using System.Threading;

namespace SharpLayout
{
    public class FontSlot
    {
        private static int _nextId;
        private readonly ConcurrentDictionary<string, FontFamilyInfo> _cache = new();

        public int Identifier { get; } = Interlocked.Increment(ref _nextId);
        
        public FontFamilyInfo FamilyInfo(string familyName)
        {
            // Fast path (zero allocations, compatible with netstandard2.0)
            if (_cache.TryGetValue(familyName, out var info))
                return info;

            return _cache.GetOrAdd(familyName, name => new FontFamilyInfo(
                this, 
                originalName: name, 
                name: name + Identifier));
        }
    }
}