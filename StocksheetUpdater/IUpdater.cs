using System.Threading.Tasks;

namespace StocksheetUpdater
{
    internal interface IUpdater
    {
        Task Run();
    }
}