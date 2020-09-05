using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CoreHtmlToImage;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using Spire.Xls;

namespace TestProjectService
{
    public class Logic
    {
        public async void DoStuff(IFormFile excelFile)
        {
            var excelInput = ParseExcel(excelFile);
            var htmlFilePath = GenerateHtml(excelInput);
            //var htmlFilePath = "Temp/3c2ea1e8.html";
            SaveHtmlAsImage(htmlFilePath);
        }

        private string SaveHtmlAsImage(string htmlFilePath)
        {
            var filename = Path.GetFileNameWithoutExtension(htmlFilePath) + ".jpg";
            var filePath = Path.Combine("Temp", filename);
            
            var converter = new HtmlConverter();
            var html = File.ReadAllText(htmlFilePath);
            var bytes = converter.FromHtmlString(html);
            File.WriteAllBytes(filePath, bytes);

            return filePath;
        }

        private List<ExcelInputRow> ParseExcel(IFormFile excelFile)
        {
            var workbook = new Workbook();
            workbook.LoadFromStream(excelFile.OpenReadStream());
            var sheet = workbook.Worksheets[0];

            var excelRows = new List<ExcelInputRow>();
            for (int i = 1; i < sheet.LastRow; i++)
            {
                var cells = sheet.Range.Rows[i].CellList;
                string productName = cells[1].Value;
                if (!string.IsNullOrEmpty(productName))
                {
                    var quantity = cells[2].Value;
                    var priceCoop365 = cells[3].NumberValue;
                    var priceNetto = cells[4].NumberValue;
                    var priceRema = cells[5].NumberValue;
                    excelRows.Add(new ExcelInputRow(productName, quantity, priceCoop365, priceNetto, priceRema));
                }
            }

            return excelRows;
        }

        private string GenerateHtml(List<ExcelInputRow> itemRows)
        {
            var mainTemplateHtml = File.ReadAllText("Assets/main-template.html");
            var itemTemplateHtml = File.ReadAllText("Assets/item-template.html");

            // generate the html content
            var itemsHtml = "";
            foreach (var itemRow in itemRows)
            {
                itemsHtml += itemTemplateHtml.Replace("{{product}}", itemRow.Name)
                    .Replace("{{price}}", itemRow.PriceCoop365.ToString("N2"));
            }
            var allHtml = mainTemplateHtml.Replace("{{items}}", itemsHtml)
                .Replace("{{logo}}", "https://eu003.kimbinocdn.com/dk/data/12/logo.png")
                .Replace("{{itemCount}}", itemRows.Count.ToString())
                .Replace("{{totalPrice}}", itemRows.Sum(x => x.PriceCoop365).ToString("N2"));

            var filename = Guid.NewGuid().ToString("N").Substring(0, 8) + ".html";
            var filePath = Path.Combine("Temp", filename);
            File.WriteAllText(filePath, allHtml);

            return filePath;
        }

        private async Task<string> SaveExcelToDisk(IFormFile excelFile)
        {
            var ext = Path.GetExtension(excelFile.FileName);
            var filename = Guid.NewGuid().ToString("N").Substring(0, 8) + ext;
            var filePath = Path.Combine("Temp", filename);
            await using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await excelFile.CopyToAsync(fileStream);
            }
            return filePath;
        }

    }

    public class ExcelInputRow
    {
        public string Name { get; }
        public double PriceCoop365 { get; }
        public double PriceNetto { get; }
        public double PriceRema { get; }

        public ExcelInputRow(string productName, string quantity, double priceCoop365, double priceNetto, double priceRema)
        {
            Name = $"{productName} {quantity}";
            PriceCoop365 = priceCoop365;
            PriceNetto = priceNetto;
            PriceRema = priceRema;
        }
    }
}
