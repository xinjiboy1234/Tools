using ExcelTools;
using ExcelTools.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper
{
    public enum Status
    {
        [ExcelOptionItemDisplay("使用")]
        USE = 1,
        [ExcelOptionItemDisplay("不使用")]
        UNUSE
    }
}
