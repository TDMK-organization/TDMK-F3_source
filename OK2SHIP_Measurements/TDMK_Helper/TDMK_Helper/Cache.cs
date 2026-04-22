using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Caching;

namespace TDMK_Helper
{
    internal class Cache
    {
        private static MemoryCache _cache = MemoryCache.Default;
        public static void Set(string key, object value, int secondsToExpire)
        {
            var policy = new CacheItemPolicy
            {
                AbsoluteExpiration = DateTimeOffset.Now.AddSeconds(secondsToExpire)
            };
            _cache.Set(key, value, policy);
        }
        public static object Get(string key)
        {
            return _cache.Get(key);
        }
        public static void Remove(string key)
        {
            _cache.Remove(key);
        }
        public static void RemoveAll()
        {
            foreach (var item in _cache)
            {
                _cache.Remove(item.Key);
            }
        }
        public static bool Contains(string key)
        {
            return _cache.Contains(key);
        }
    }
}
