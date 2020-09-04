using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace TestProjectService
{
    public class Logic
    {
        public async void DoStuff(IFormFile excelFile)
        {
            var filePath = await SaveExcelFileToDisk(excelFile);
            var excelInput = ParseExcelInput(filePath);
        }

        private ExcelInput ParseExcelInput(string filePath)
        {

            using var package = new ExcelPackage(new FileInfo(filePath));
            var sheet = package.Workbook.Worksheets[0];
            var c0 = sheet.Cells[0, 0];

            return new ExcelInput();
        }

        private async Task<string> SaveExcelFileToDisk(IFormFile excelFile)
        {
            var ext = Path.GetExtension(excelFile.FileName);
            var filename = Guid.NewGuid().ToString("N") + ext;
            var filePath = Path.Combine("Temp", filename);
            Directory.CreateDirectory(filePath);

            await using var fileStream = new FileStream(filePath, FileMode.Create);
            await excelFile.CopyToAsync(fileStream);

            return filePath;
        }

    }

    public class ExcelInput
    {
        public string ImagePath;
        public List<ExcelInputRow> ProductRows = new List<ExcelInputRow>();
    }

    public class ExcelInputRow
    {
        public string Name { get; }
        public float PriceCoop365 { get; }
        public float PriceNetto { get; }
        public float PriceRema { get; }

        public ExcelInputRow(string name, float priceCoop365, float priceNetto, float priceRema)
        {
            Name = name;
            PriceCoop365 = priceCoop365;
            PriceNetto = priceNetto;
            PriceRema = priceRema;
        }
    }
}
