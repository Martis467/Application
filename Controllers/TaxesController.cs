using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using TaxManager.DAL;
using TaxManager.Models;
using TaxManager.Models.Database;
using TaxManager.Services;

namespace TaxManager.Controllers
{
    [Route("api/taxes")]
    public class TaxesController : Controller
    {
        private readonly TaxContext _context;
        private readonly IMunicipalityTaxService _taxService;

        public TaxesController(TaxContext context)
        {
            _context = context;
            _taxService = new MunicipalityTaxService(_context);
        }

        [HttpGet]
        public async Task<IActionResult> Details(string municpality, DateTime? fromTime, DateTime? toTime)
        {
            try
            {
                return Ok(await _taxService.GetByMunicipalityAndDate(municpality, fromTime, toTime));
            }
            catch
            {
                return NotFound();
            }
        }

        [HttpPost]
        public async Task<IActionResult> AddMunicipalityTax([FromBody] MunicipalityTax tax)
        {
            try
            {
                tax.Verify();
                return Ok(await _taxService.AddMunicipalityTaxes(tax));
            }
            catch (Exception e)
            {
                return BadRequest();
            }
        }

        [HttpPost("import")]
        public async Task<IActionResult> ImportMunicipalities(IFormFile file)
        {
            try
            {
                if (file == null)
                    return BadRequest();

                if (!file.FileName.EndsWith(".xlsx"))
                    return BadRequest();

                await _taxService.ImportMunicipalities(file.OpenReadStream());
                return Ok();
            }
            catch(Exception e)
            {
                return BadRequest();
            }
        }
    }
}
