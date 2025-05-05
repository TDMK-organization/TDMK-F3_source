using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public static class DictionaryService
    {
        public static void AddOrUpdate<TValue>(IDictionary<string, TValue> dic, string key, TValue value)
        {
            if (dic.ContainsKey(key))
            {
                dic[key] = value;
            }
            else
            {
                dic.Add(key, value);
            }
        }
        public static string GetValueOrNull(IDictionary<string, string> dic, string key)
        {
            if(dic.TryGetValue(key, out string s))
            {
                return s;
            }
            return s;
        }
        public static IDictionary<TKey, TValue> MergeDictionaries<TKey, TValue>(IDictionary<TKey, TValue> dict1, IDictionary<TKey, TValue> dict2)
        {
            if (dict1 == null)
            {
                return dict2 != null ? new Dictionary<TKey, TValue>(dict2) : new Dictionary<TKey, TValue>();
            }

            if (dict2 == null)
            {
                return new Dictionary<TKey, TValue>(dict1);
            }

            Dictionary<TKey, TValue> mergedDictionary = new Dictionary<TKey, TValue>(dict1);

            foreach (var kvp in dict2)
            {
                mergedDictionary[kvp.Key] = kvp.Value;
            }

            return mergedDictionary;
        }
    }
}
