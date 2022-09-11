using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaTraderConversionLib.Dtos
{
    internal class TransactionDto
    {
        public string Type { get; set; } = "";
        public decimal Amount { get; set; }
        public string AssetName { get; set; } = "";
        public DateTime BuyTime { get; set; }
        public DateTime SellTime { get; set; }
        public decimal BuyPrice { get; set; }
        public decimal SellPrice { get; set; }
        public decimal Profit { get; set; }
    }
}
