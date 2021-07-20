using ExcelTools;
using ExcelTools.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper
{
    public class Person
    {
        [ExcelHeadDisplay("ID", 1)]
        public int Id { get; set; }
        [ExcelHeadDisplay("名称", 2)]
        public string Name { get; set; }
        [ExcelHeadDisplay("选项",4)]
        public Status Status { get; set; }
        [ExcelHeadDisplay("Quantity",5)]
        public decimal Quantity { get; set; }
        [ExcelHeadDisplay("选项1",3)]
        public Status Status1 { get; set; }
    }
}
