using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace StocksheetUpdater
{
    class Updater : IUpdater
    {
        private readonly MySettings _settings;

        private List<StockSpreadSheet> SpreadSheets { get; set; } = new List<StockSpreadSheet>();


        public Updater(IOptions<MySettings> settings)
        {
            _settings = settings.Value;
        }

        public async Task Run()
        {

            foreach (var book in _settings.Books)
            {
                Console.WriteLine($"Opening book {book.Filename}");

                foreach (string sheetName in book.SheetNames)
                {
                    Console.WriteLine($"Opening sheet {sheetName}");

                    var newSheet =
                       new StockSpreadSheet(sheetName, _settings.FilePath + @"/" + book.Filename, _settings);

                    Console.WriteLine("ajajaj");

                    SpreadSheets.Add(newSheet);
                }
            }
        }
    }
}
