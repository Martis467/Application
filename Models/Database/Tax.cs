using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using TaxManager.Extensions;

namespace TaxManager.Models.Database
{
    public class Tax
    {
        public enum TaxType
        {
            Daily = 1,
            Weekly = 2,
            Monthly = 3,
            Yearly = 4
        }
        public int Id { get; set; }
        [Required]
        public decimal Value { get; set; }
        [Required]
        public TaxType Type { get; set; }
        [Required]
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        // Foreign key
        public int MunicipalityId { get; set; }

        public static Tax FromMunicipalityTax(MunicipalityTax municipalityTax)
        {
            // This looks ugly because of my chosen exception handling
            return new Tax
            {
                Value = municipalityTax.Value,
                Type = municipalityTax.Type.Value,
                StartDate = municipalityTax.StartDate.Value,
                EndDate = municipalityTax.StartDate.Value.EndingDate(municipalityTax.Type.Value)
            };
        }
    }
}