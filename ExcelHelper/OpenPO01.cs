using ExcelTools.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper
{
    public class OpenPO01
    {
        [ExcelHeadDisplay("Item$SV$Xxinv_Item_Number")]
        public string ItemNumber { get; set; }
        [ExcelHeadDisplay("PO Number")]
        public string PoNumber { get; set; }
    }
}
