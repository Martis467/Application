using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TaxManager.Extensions
{
    public static class StringExtensions
    {
        public static T StringToEnum<T>(this string str)
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();

            return (T)Enum.Parse(typeof(T), str, true);
        }
    }
}
