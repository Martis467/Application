using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using TaxManager.Models.Database;

namespace TaxManager.DAL
{
    public static class DataInitializer
    {
        public static void Initialize(TaxContext context)
        {
            context.Database.EnsureCreated();

            if (context.Municapilities.Any())
                return;

            var municipality = new Municipality
            {
                Name = "Vilnius"
            };

            context.Municapilities.Add(municipality);
            context.SaveChanges();

           var taxes = new List<Tax>()
            {
                new Tax
                {
                    MunicipalityId = 1,
                    Type = Tax.TaxType.Yearly,
                    StartDate = new DateTime(2016, 1, 1),
                    EndDate = new DateTime(2016, 1, 1).AddYears(1).AddDays(-1),
                    Value = 0.2m
                },
                new Tax
                {
                    MunicipalityId = 1,
                    Type = Tax.TaxType.Monthly,
                    StartDate = new DateTime(2016, 5, 1),
                    EndDate = new DateTime(2016, 1, 1).AddMonths(1).AddDays(-1),
                    Value = 0.4m
                },
                new Tax
                {
                    MunicipalityId = 1,
                    Type = Tax.TaxType.Daily,
                    StartDate = new DateTime(2016, 1, 1),
                    EndDate = new DateTime(2016, 1, 1),
                    Value = 0.1m
                },
                new Tax
                {
                    MunicipalityId = 1,
                    Type = Tax.TaxType.Daily,
                    StartDate = new DateTime(2016, 12, 25),
                    EndDate = new DateTime(2016, 12, 25),
                    Value = 0.1m
                }
            };

            taxes.ForEach(t => context.Taxes.Add(t));
            context.SaveChanges();
        }
    }
}