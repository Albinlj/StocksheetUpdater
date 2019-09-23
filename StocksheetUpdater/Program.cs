using System;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Extensions.Configuration.FileExtensions;
using Microsoft.Extensions.DependencyInjection;

namespace StocksheetUpdater
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                   .SetBasePath(Utils.GetApplicationRoot())
                   .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            IConfiguration configuration = builder.Build();

            var serviceProvider = new ServiceCollection()
                .AddOptions()
                .Configure<MySettings>(configuration.GetSection("MySettings"))
                .AddSingleton<IUpdater, Updater>()
                .BuildServiceProvider();

            var updater = serviceProvider.GetService<IUpdater>();
            await updater.Run();
        }
    }
}
