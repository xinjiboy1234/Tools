﻿using ExcelTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelHelper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var et = new ExcelHelper<Person>();
            var p = new Person
            {
                Id = 123,
                Status = Status.USE,
                Status1 = Status.UNUSE,
                Name = "ddd",
                Quantity = (decimal)123.12
            };
            // et.SaveExcelFromCollection(new List<Person>{p}, $@"{AppDomain.CurrentDomain.BaseDirectory}\2.xlsx");

            var dataListByExcelPath = et.GetDataListByExcelPath($@"{AppDomain.CurrentDomain.BaseDirectory}2.xlsx");
        }
    }
}
