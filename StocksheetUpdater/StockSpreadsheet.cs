using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;
using OfficeOpenXml;

namespace StocksheetUpdater
{
    class StockSpreadSheet : IDisposable
    {
        public string SheetName { get; set; }
        public string Symbol { get; set; }
        public DateTime LastUpdated { get; set; }
        public MySettings Settings { get; }
        public ExcelPackage Package { get; set; }
        public ExcelWorksheet Worksheet { get; set; }

        public StockSpreadSheet(string SheetName, string filePath, MySettings settings)
        {
            Settings = settings;
            Package = new ExcelPackage(new FileInfo(filePath));
            Worksheet = Package.Workbook.Worksheets[SheetName];

            ReadSymbol();
            ReadLastUpdateDate();
            Dispose();
        }

        public void AddRow(ToRow row)
        {

        }

        public void ReadSymbol()
        {

        }

        private void ReadLastUpdateDate()
        {
            var firstDateAddressCell = Worksheet.Cells[Settings.FirstDateAddress].Start;

            LastUpdated = Worksheet.Cells[firstDateAddressCell.Address + ":" + ExcelCellAddress.GetColumnLetter(firstDateAddressCell.Column)]
                .Where(c => c.Value != null)
                .LastOrDefault()
                .GetValue<DateTime>();

            Console.WriteLine($"Last updated {LastUpdated.ToLongDateString()}");

        }

        private static async Task UpdateSheet()
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(
                    "https://www.alphavantage.co/query?apikey=R0HDGNWY50BPN6N6&function=TIME_SERIES_DAILY&symbol=MSFT");

                Console.WriteLine(response.Content.Headers.ToString());
            }
        }

        public void Print()
        {
        }

        public void Dispose()
        {
            Package.Dispose();
        }
    }
}
