using System;
using System.Collections.Generic;
using System.Linq;

namespace CalendarSync
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<IEnumerable<T>> Chunk<T>(this IEnumerable<T> source, int chunk_size)
        {
            while (source.Any())
            {
                yield return source.Take(chunk_size);
                source = source.Skip(chunk_size);
            }
        }

        public static void Apply<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (var item in source)
            {
                action(item);
            }
        }
    }
}