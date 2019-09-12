using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace StocksheetUpdater
{
    class StockSpreadSheet
    {
        private readonly string _filePath;

        public string Name { get; set; }
        public string Symbol { get; set; }
        public DateTime LastUpdated { get; set; }
        public int LastDateRow { get; set; }

        public StockSpreadSheet(string filePath)
        {
            LastDateRow = 11;
            _filePath = filePath;
        }

        public void AddRow(ToRow row)
        {

        }

        public async Task ReadName()
        {
            FileInfo existingFile = new FileInfo(_filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                const string dateColumn = "AB";
                const string nameAdress = "AB8";

                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                Name = worksheet.Cells[nameAdress].Value.ToString();

                HttpClient client = new HttpClient();
                var response = await client.GetAsync(
                    "https://www.alphavantage.co/query?apikey=R0HDGNWY50BPN6N6&function=TIME_SERIES_DAILY&symbol=MSFT");

                Console.WriteLine(response.Content.Headers.ToString());
                Console.WriteLine("strut");

                while (worksheet.Cells[dateColumn + (LastDateRow + 1)].Value != null)
                {
                    LastDateRow++;
                }

                LastUpdated = DateTime.Parse(worksheet.Cells["AB" + LastDateRow].Value.ToString());
            }
        }

        public void Print()
        {
            Console.WriteLine(Name);
            Console.WriteLine(LastUpdated.Date);
        }
    }
}
