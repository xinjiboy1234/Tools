using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTools.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelOptionItemDisplayAttribute : Attribute
    {
        public string OptionDisplay { get; set; }

        public ExcelOptionItemDisplayAttribute(string optionDisplay)
        {
            OptionDisplay = optionDisplay;
        }
    }
}
