using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;

namespace ExcelTools
{
    public static class ExcelStyleTools
    {
        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="color"></param>
        /// <param name="borderStyle"></param>
        public static void SetBorder(ref ExcelRange cell, Color color,
            ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            cell.Style.Border.BorderAround(borderStyle, color);
        }

        /// <summary>
        /// 设置字体大小
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="fontSize"></param>
        public static void SetFontSize(ref ExcelRange cell, int fontSize = 12)
        {
            cell.Style.Font.Size = fontSize;
        }

        /// <summary>
        /// 设置字体
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="fontName"></param>
        public static void SetFont(ref ExcelRange cell, string fontName)
        {
            cell.Style.Font.Name = fontName;
        }

        /// <summary>
        /// 设置字体颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="color"></param>
        public static void SetFontColor(ref ExcelRange cell, Color color)
        {
            cell.Style.Font.Color.SetColor(color);
        }

        /// <summary>
        /// 设置字体加粗
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="boldFlag"></param>
        public static void SetFontBold(ref ExcelRange cell, bool boldFlag)
        {
            cell.Style.Font.Bold = true;
        }

        /// <summary>
        /// 设置背景颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="color"></param>
        public static void SetBackground(ref ExcelRange cell, Color color)
        {
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// 设置对齐
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="excelHorizontalAlignment"></param>
        /// <param name="excelVerticalAlignment"></param>
        public static void SetAlign(ref ExcelRange cell,
            ExcelHorizontalAlignment excelHorizontalAlignment = ExcelHorizontalAlignment.Center,
            ExcelVerticalAlignment excelVerticalAlignment = ExcelVerticalAlignment.Center)
        {
            cell.Style.HorizontalAlignment = excelHorizontalAlignment;
            cell.Style.VerticalAlignment = excelVerticalAlignment;
        }
    }
}