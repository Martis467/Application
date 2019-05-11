using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using static TaxManager.Models.Database.Tax;

namespace TaxManager.Extensions
{
    public static class DateTimeExtensions
    {
        public static DateTime EndingDate(this DateTime dateTime, TaxType type)
        {
            switch (type)
            {
                case TaxType.Daily:
                    return dateTime;
                case TaxType.Weekly:
                    return dateTime.AddDays(7);
                case TaxType.Monthly:
                    return dateTime.AddMonths(1).AddDays(-1);
                case TaxType.Yearly:
                    return dateTime.AddYears(1).AddDays(-1);
            }

            // Should not happen though
            return dateTime;
        }
    }
}