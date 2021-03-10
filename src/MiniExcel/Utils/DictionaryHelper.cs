namespace MiniExcelLibs.Utils
{
    using System;
    using System.Collections.Generic;

    internal static class DictionaryHelper
    {
        public static TValue GetValueOrDefault<TKey, TValue>
        (this IDictionary<TKey, TValue> dictionary,
         TKey key,
         TValue defaultValue)
        {
            TValue value;
            return dictionary.TryGetValue(key, out value) ? value : defaultValue;
        }

        public static TValue GetValueOrDefault<TKey, TValue>
             (this IDictionary<TKey, TValue> dictionary,
              TKey key,
              Func<TValue> defaultValueProvider)
        {
            TValue value;
            return dictionary.TryGetValue(key, out value) ? value
                  : defaultValueProvider();
        }
    }
}
