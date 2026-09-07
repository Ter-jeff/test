using System.Collections.Generic;
using System.Linq;

namespace CommonLib.Extension
{
    public static class CollectionExtensions
    {
        public static void AddRange<T>(this ICollection<T> tCollect, IEnumerable<T> ts)
        {
            foreach (T item in ts)
            {
                tCollect.Add(item);
            }
        }

        public static IEnumerable<IEnumerable<T>> CartesianProduct<T>(this IEnumerable<IEnumerable<T>> sequences)
        {
            IEnumerable<IEnumerable<T>> emptyProduct = [[]];
            return sequences.Aggregate(
                emptyProduct,
                (accumulator, sequence) =>
                    from acc in accumulator
                    from item in sequence
                    select acc.Concat([item])
            );
        }
    }
}
