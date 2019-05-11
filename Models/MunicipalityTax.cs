using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using Microsoft.EntityFrameworkCore;
using TaxManager.Models.Database;
using static TaxManager.Models.Database.Tax;

namespace TaxManager.Models
{
    public class MunicipalityTax
    {
        public decimal Value { get; set; }
        public TaxType? Type { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string MunicipalityName { get; set; }

        public MunicipalityTax() { }
        public MunicipalityTax(Tax t, Municipality m)
        {
            Value = t.Value;
            Type = t.Type;
            StartDate = t.StartDate;
            EndDate = t.EndDate;
            MunicipalityName = m.Name;
        }

        public void Verify()
        {
            if ( Value <= 0 || !Type.HasValue || !StartDate.HasValue || MunicipalityName == null)
                throw new Exception();
        }

        public static IEnumerable<MunicipalityTax> FromTaxesAndMunicipalities(List<Tax> taxes, List<Municipality> municipalities)
        {
            var list = new List<MunicipalityTax>();
            municipalities.ForEach(m =>
            taxes.ForEach(t =>
            {
                if (t.MunicipalityId == m.Id)
                    list.Add(new MunicipalityTax(t, m));
            }));

            return list;
        }
    }
}