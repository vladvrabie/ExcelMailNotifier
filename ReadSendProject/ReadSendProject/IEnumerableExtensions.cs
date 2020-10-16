using System.Collections.Generic;
using System.Linq;

namespace ReadSendProject
{
    public static class IEnumerableExtensions
    {
        /// <summary>
        /// Credits here: https://stackoverflow.com/a/52796682
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <returns></returns>
        public static bool IsNullOrEmpty<T>(this IEnumerable<T> source)
        {
            return source == null || !source.Any();
        }
    }
}
