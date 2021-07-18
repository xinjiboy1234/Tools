using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelTools
{
    public class ExcelHelper<T> where T : class, new()
    {
        public ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        /// <summary>
        /// 将Excel数据转换成对应类型的列表
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public IEnumerable<T> GetDataListByExcelPath(string path)
        {
            var fi = new FileInfo(path);
            using var excelPackage = new ExcelPackage(fi);
            var workSheet = excelPackage.Workbook?.Worksheets?.FirstOrDefault();
            if (workSheet == null) return null;
            var colIndexDictionary = GetColumnIndexAndPropertiesMapping(workSheet);
            var list = new List<T>();
            for (var i = 2; i <= workSheet.Dimension.End.Row; i++)
            {
                var t = new T();
                var props = t.GetType().GetProperties();
                foreach (var item in colIndexDictionary)
                {
                    var prop = props.FirstOrDefault(x => x.Name == item.Value.Name);
                    ObjectUtils.FillValueToProperty(ref prop, t, workSheet.Cells[i, item.Key].Value);
                }

                list.Add(t);
            }
            
            return list;
        }

        /// <summary>
        /// 获取excel表头顺序和与其对应类的属性键值对
        /// </summary>
        /// <param name="excelWorksheet"></param>
        /// <returns></returns>
        private Dictionary<int, PropertyInfo> GetColumnIndexAndPropertiesMapping(ExcelWorksheet excelWorksheet)
        {
            var colIndexDictionary = new Dictionary<int, PropertyInfo>();
            for (var i = excelWorksheet.Dimension.Start.Column; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                colIndexDictionary[i] = ObjectUtils.GetPropertyInfoByAttributeDescription<T>(excelWorksheet.Cells[1, i].Value.ToString());
            }

            return colIndexDictionary;
        }
    }
}
