using System;
using System.Collections.Generic;
using System.Text;

namespace StocksheetUpdater
{
    class MySettings
    {
        public string FirstDateAddress { get; set; }
        public int IndexOfFirstWorksheet { get; set; }
        public string SymbolAddress { get; set; }
        public string FilePath { get; set; }
        public Book[] Books { get; set; }
    }
}
