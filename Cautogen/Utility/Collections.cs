using System.Collections.Generic;
using System.Linq;

namespace Cautogen.Utility
{
    public static class Collections
    {
        public static List<List<T>> Zip<T>(IEnumerable<IEnumerable<T>> items)
        {
            var result = new List<List<T>>();
            var enumerators = items.Select(item => item.GetEnumerator()).ToList();
            IEnumerable<bool> hasNexts = from enumerator in enumerators select enumerator.MoveNext();

            while (hasNexts.Aggregate((a, b) => a && b))
            {
                result.Add(enumerators.Select(enumerator => enumerator.Current).ToList());
                hasNexts = from enumerator in enumerators select enumerator.MoveNext();
            }
            return result;
        }
    }
}
