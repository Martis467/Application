using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using TaxManager.DAL;
using TaxManager.Extensions;
using TaxManager.Models;
using TaxManager.Models.Database;

namespace TaxManager.Services
{
    public interface IMunicipalityTaxService
    {
        Task<MunicipalityTax> AddMunicipalityTaxes(MunicipalityTax municipalityTax, bool saveToDb = false);
        Task<IEnumerable<MunicipalityTax>> GetByMunicipalityAndDate(string municipality, DateTime? fromTime, DateTime? toTime);
        Task ImportMunicipalities(Stream stream);
    }
    public class MunicipalityTaxService : IMunicipalityTaxService
    {
        private readonly TaxContext _context;

        public MunicipalityTaxService(TaxContext context)
        {
            _context = context;
        }

        public async Task<IEnumerable<MunicipalityTax>> GetByMunicipalityAndDate(string municipality, DateTime? fromTime, DateTime? toTime)
        { 
            // Recreating to keep only date
            fromTime = new DateTime(fromTime.Value.Year, fromTime.Value.Month, fromTime.Value.Day);
            toTime = new DateTime(toTime.Value.Year, toTime.Value.Month, toTime.Value.Day);

            var mun = _context.Municapilities.FirstOrDefault(m => m.Name == municipality);

            if (mun == null)
                throw new Exception();

            var taxes = await _context.Taxes.Where(t => t.MunicipalityId == mun.Id).ToListAsync();

            // Not sure if this is a good way to apply filters, should be done in sql querry
            if (fromTime != DateTime.MinValue)
                taxes = taxes.Where(t => t.StartDate >= fromTime).ToList();

            if(toTime != DateTime.MaxValue)
                taxes = taxes.Where(t => t.EndDate <= toTime).ToList();

            return taxes.Select(t => new MunicipalityTax(t, mun));
        }

        public async Task<MunicipalityTax> AddMunicipalityTaxes(MunicipalityTax municipalityTax, bool saveToDb = false)
        {
            var municipality = _context.Municapilities.Where(m => m.Name == municipalityTax.MunicipalityName).FirstOrDefault();
            var tax = Tax.FromMunicipalityTax(municipalityTax);

            if (municipality == null)
            {
                _context.Municapilities.Add(new Municipality { Name = municipalityTax.MunicipalityName });
                _context.SaveChanges();
                municipality = _context.Municapilities.Where(m => m.Name == municipalityTax.MunicipalityName).FirstOrDefault();
            }

            tax.MunicipalityId = municipality.Id;

            if (!TaxIsValid(tax))
                throw new Exception("Invalid tax");

            _context.Taxes.Add(tax);
            await _context.SaveChangesAsync();

            municipalityTax.EndDate = tax.EndDate;

            return municipalityTax;
        }

        public async Task ImportMunicipalities(Stream stream)
        {
            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets[0];

                ParseHeader(ref workSheet);
                var taxes = ParseWorksheet(ref workSheet);

                // Add new taxes
                foreach (var tax in taxes)
                    await AddMunicipalityTaxes(tax, true);
            }
        }

        private IEnumerable<MunicipalityTax> ParseWorksheet(ref ExcelWorksheet workSheet)
        {
            var row = workSheet.Dimension.Start.Row + 1;
            var list = new List<MunicipalityTax>();

            while(RowHasMandatoryValues(ref workSheet, row))
            {
                list.Add(ParseTax(ref workSheet, row));
                row++;
            }

            return list;
        }

        private MunicipalityTax ParseTax(ref ExcelWorksheet workSheet, int row)
        {
            var tax = new MunicipalityTax();

            tax.MunicipalityName = workSheet.Cells[row, 1].Value.ToString().Trim();
            tax.Value = ParseValue(workSheet.Cells[row, 2].Value.ToString());
            tax.Type = workSheet.Cells[row, 3].Value.ToString().StringToEnum<Tax.TaxType>();
            tax.StartDate = Convert.ToDateTime(workSheet.Cells[row, 4].Value.ToString());

            return tax;
        }

        private bool RowHasMandatoryValues(ref ExcelWorksheet workSheet, int row)
        {
            // If we find an empyt row we can asume we reached the end
            if (workSheet.Cells[row, 1, row, 4].All(cell => cell.Value == null))
                return false;

            return true;
        }

        private void ParseHeader(ref ExcelWorksheet workSheet)
        {
            // Validate if predefined headers are correct

            if (!workSheet.Cells[1, 1].Value.ToString().Equals("Municipality") ||
                !workSheet.Cells[1, 2].Value.ToString().Equals("Tax value") ||
                !workSheet.Cells[1, 3].Value.ToString().Equals("Schedule") ||
                !workSheet.Cells[1, 4].Value.ToString().Equals("Starting date"))
                throw new Exception();
        }

        // Utility function to parse decimal string for cultural invariance
        private decimal ParseValue(string value)
        {
            value = value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator);
            var valueFormat = decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue);

            if (!valueFormat)
                throw new Exception();

            return parsedValue;
        }

        private bool TaxIsValid(Tax tax)
        {
            var taxes = _context.Taxes.Where(t => t.MunicipalityId == tax.MunicipalityId).ToList();

            // Ading another yearly doesn't seem right
            if (tax.Type == Tax.TaxType.Yearly && taxes.FirstOrDefault(t => t.Type == Tax.TaxType.Yearly) != null)
                return false;

            // Monthly taxes are bit trickier, not sure if the starting date always starts at the first day of the month
            // So I'll be adding a general check
            // Same goes for weekly
            if (taxes.Any(t => t.StartDate == tax.StartDate && t.EndDate == tax.EndDate))
                return false;

            return true;
        }
    }
}