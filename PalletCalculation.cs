using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderPalletCount
{
    public class PalletCalculation
    {
        public string EventID { get; set; }
        public string SKU { get; set; }
        public string Description { get; set; }
        public string CartonWeight { get; set; }
        public string OrderQuantity { get; set; }
        public string CartonQty { get; set; }
        public string SKUID { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
    }
}
