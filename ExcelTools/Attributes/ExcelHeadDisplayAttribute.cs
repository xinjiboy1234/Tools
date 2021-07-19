using System;

namespace ExcelTools.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelHeadDisplayAttribute : Attribute
    {
        public string HeadDisplay { get; set; }
        public ExcelHeadDisplayAttribute(string head)
        {
            HeadDisplay = head;
        }
    }
}
