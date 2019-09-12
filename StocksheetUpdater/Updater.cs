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
        private readonly HttpClient _client;
        private readonly IOptions<MySettings> _settings;


        public Updater(IOptions<MySettings> settings)
        {
            _settings = settings;
        }

        public async Task Run()
        {
            Console.WriteLine(_settings.Value.Laser);


            Console.WriteLine("KOSAD!");



            StockSpreadSheet sheet =
                new StockSpreadSheet(@"C:\Users\Alb\source\repos\StocksheetUpdater\StocksheetUpdater\Bok_A.xlsx");
            await sheet.ReadName();
            sheet.Print();


        }


    }
}
