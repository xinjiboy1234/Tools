using System;

namespace ExcelTools.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelHeadDisplayAttribute : Attribute
    {
        public string HeadDisplay { get; }
        public int Index { get;}
        public ExcelHeadDisplayAttribute(string head)
        {
            HeadDisplay = head;
        }

        public ExcelHeadDisplayAttribute(string headDisplay, int index)
        {
            HeadDisplay = headDisplay;
            Index = index;
        }
    }
}
