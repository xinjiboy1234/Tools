using ExcelTools;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper
{
    public class Person
    {
        [Head("ID")]
        public int Id { get; set; }
        [Head("名称")]
        public string Name { get; set; }
        [Head("选项")]
        public Status Status { get; set; }
    }
}
