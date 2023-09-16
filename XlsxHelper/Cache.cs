using System.Buffers;

namespace XlsxHelper;

internal class Cache<TKey, TValue> where TKey : notnull
{
    private readonly Dictionary<TKey, CacheEntry<TValue>> _cache;
    private readonly Func<TKey, TValue> _store;
    private readonly int _cacheSize;

    private int ops;
    public Cache(Func<TKey, TValue> store, int cacheSize)
    {
        _store = store;
        _cacheSize = cacheSize;
        _cache = new Dictionary<TKey, CacheEntry<TValue>>(cacheSize);
    }

    public TValue GetOrAddValue(TKey key)
    {
        if (Get(key, out var value))
            return value!;

        RemoveUnusedKeys();

        value = _store(key);
        Add(key, value);
        return value;
    }

    private bool Get(TKey key, out TValue? value)
    {
        ops++;
        if (_cache.TryGetValue(key, out var cacheEntry))
        {
            cacheEntry.IncrementHit();
            value = cacheEntry.Value;
            return true;
        }
        value = default;
        return false;
    }

    private void Add(TKey key, TValue value)
    {
        if (_cache.Count < _cacheSize)
            _cache.Add(key, new CacheEntry<TValue>(value));
    }

    private void RemoveUnusedKeys()
    {
        var zeroHitCount = 0;
        if (ops >= _cacheSize * 10)
        {
            foreach (var ce in _cache.Values)
            {
                ce.DecrementHit(2);
                if (ce.HitsCount <= 0)
                    zeroHitCount++;
            }
            ops = 0;
        }

        if (zeroHitCount > 0)
        {
            var i = 0;
            var keysToRemove = ArrayPool<TKey>.Shared.Rent(zeroHitCount);
            foreach (var ce in _cache)
            {
                if (ce.Value.HitsCount <= 0)
                    keysToRemove[i++] = ce.Key;
            }

            for (i = 0; i < zeroHitCount; i++)
            {
                _cache.Remove(keysToRemove[i]);
            }
            ArrayPool<TKey>.Shared.Return(keysToRemove);
        }
    }

    private class CacheEntry<T>
    {
        public T Value { get; }
        public int HitsCount { get; private set; }

        public CacheEntry(T value)
        {
            Value = value;
        }

        internal void IncrementHit()
        {
            HitsCount++;
        }

        internal void DecrementHit(int count)
        {
            HitsCount-= count;
        }
    }
}
