using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CompareExcelItem.Model
{
    public class Material
    {
        public string SerialNumber { get; set; }

        public string Model { get; set; }
        public string Description { get; set; }
        public string Quantity { get; set; }
        public string Remark { get; set; }
    }
}