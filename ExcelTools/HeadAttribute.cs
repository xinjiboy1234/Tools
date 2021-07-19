using System;

namespace ExcelTools
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class HeadAttribute : Attribute
    {
        public string Head { get; set; }
        public HeadAttribute(string head)
        {
            Head = head;
        }
    }
}
