using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using WebApplication2.Extensions;
using WebApplication2.Models;

namespace WebApplication2.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PessoaController : ControllerBase
    {
        private readonly List<Pessoa> _pessoas = new()
        {
            new Pessoa { Id = 0, Nome = "Gabriel", Sobrenome = "Santana" },
            new Pessoa { Id = 1, Nome = "Izadora", Sobrenome = "Lee" },
            new Pessoa { Id = 2, Nome = "Rosivania", Sobrenome = "Santana" }
        };

        private readonly List<WeatherForecast> _weathers = new()
        {
            new WeatherForecast { Date = DateTime.Now, Summary = "Item 1", TemperatureC = 30 },
            new WeatherForecast { Date = DateTime.Now.AddDays(1), Summary = "Item 2", TemperatureC = 50 },
            new WeatherForecast { Date = DateTime.Now.AddDays(2), Summary = "Item 3", TemperatureC = 54 }
        };

        [HttpGet]
        public IActionResult Get()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo("MyWorkbook.xlsx")))
            {
                // Tabs
                var equity = package.Workbook.Worksheets.Add("Equity").GenerateWorksheet(_pessoas);
                var coveredCallWarrant = package.Workbook.Worksheets.Add("Covered Call Warrant").GenerateWorksheet(_weathers);
                var equityOptions = package.Workbook.Worksheets.Add("Equity Options").GenerateWorksheet(_pessoas);
                var risk = package.Workbook.Worksheets.Add("Risk by Position Types").GenerateWorksheet(_weathers);
                var collateralPosition = package.Workbook.Worksheets.Add("Collateral Position").GenerateWorksheet(_pessoas);
                var marginSumm = package.Workbook.Worksheets.Add("Margin Summary").GenerateWorksheet(_weathers);

                var byteArray = package.GetAsByteArray();

                return File(byteArray, "application/excel", "MyWorkbook.xlsx");
            }
        }

        [HttpPost]
        public IActionResult Post([FromBody] Pessoa p)
        {
            _pessoas.Add(p);
            return Ok(_pessoas);
        }

        [HttpGet("{id}")]
        public IActionResult Get(int id)
        {
            var pessoa = _pessoas.Where(p => p.Id == id).FirstOrDefault();
            return Ok(pessoa);
        }
    }
}
