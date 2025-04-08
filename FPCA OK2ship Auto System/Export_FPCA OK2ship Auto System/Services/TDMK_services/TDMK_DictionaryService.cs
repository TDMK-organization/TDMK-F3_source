using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public static class TDMK_DictionaryService
    {
        
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
