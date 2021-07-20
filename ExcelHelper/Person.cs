using ExcelTools;
using ExcelTools.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper
{
    public class Person
    {
        [ExcelHeadDisplay("ID")]
        public int Id { get; set; }
        [ExcelHeadDisplay("名称")]
        public string Name { get; set; }
        [ExcelHeadDisplay("选项")]
        public Status Status { get; set; }
        [ExcelHeadDisplay("Quantity")]
        public decimal Quantity { get; set; }
        [ExcelHeadDisplay("选项1")]
        public Status Status1 { get; set; }
    }
}
