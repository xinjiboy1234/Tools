using ExcelTools.Attributes;
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
            // 存储枚举类成员和描述对应的键值对集合
            var dictionaryOfEnumMemberAndDescription = new Dictionary<string, Dictionary<string, object>>();
            // 根据Excel头部与类的标注信息想对应查找其属性，并存储与键值对中
            var colIndexDictionary = GetColumnIndexAndPropertiesMapping(workSheet, ref dictionaryOfEnumMemberAndDescription);

            var list = new List<T>();
            for (var i = 2; i <= workSheet.Dimension.End.Row; i++)
            {
                //var t = new T();
                var t = (T)Activator.CreateInstance(typeof(T));
                var props = t.GetType().GetProperties();
                foreach (var item in colIndexDictionary)
                {
                    var prop = props.FirstOrDefault(x => x.Name == item.Value.Name);
                    FillValueToProperty(ref prop, t, workSheet.Cells[i, item.Key].Value, dictionaryOfEnumMemberAndDescription);
                }

                list.Add(t);
            }

            return list;
        }

        /// <summary>
        /// 将Excel数据转换成对应类型的列表
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public IEnumerable<T> GetDataListByExcelStream(Stream stream)
        {
            using var excelPackage = new ExcelPackage(stream);
            var workSheet = excelPackage.Workbook?.Worksheets?.FirstOrDefault();
            if (workSheet == null) return null;
            // 存储枚举类成员和描述对应的键值对集合
            var dictionaryOfEnumMemberAndDescription = new Dictionary<string, Dictionary<string, object>>();
            // 根据Excel头部与类的标注信息想对应查找其属性，并存储与键值对中
            var colIndexDictionary = GetColumnIndexAndPropertiesMapping(workSheet, ref dictionaryOfEnumMemberAndDescription);

            var list = new List<T>();
            for (var i = 2; i <= workSheet.Dimension.End.Row; i++)
            {
                //var t = new T();
                var t = (T)Activator.CreateInstance(typeof(T));
                var props = t.GetType().GetProperties();
                foreach (var item in colIndexDictionary)
                {
                    var prop = props.FirstOrDefault(x => x.Name == item.Value.Name);
                    FillValueToProperty(ref prop, t, workSheet.Cells[i, item.Key].Value, dictionaryOfEnumMemberAndDescription);
                }

                list.Add(t);
            }

            return list;
        }




        /// <summary>
        /// 获取excel表头顺序和与其对应类的属性键值对
        /// </summary>
        /// <param name="excelWorksheet"></param>
        /// <param name="enumProperties"></param>
        /// <returns></returns>
        private Dictionary<int, PropertyInfo> GetColumnIndexAndPropertiesMapping(ExcelWorksheet excelWorksheet, ref Dictionary<string, Dictionary<string, object>> enumProperties)
        {
            var colIndexDictionary = new Dictionary<int, PropertyInfo>();
            for (var i = excelWorksheet.Dimension.Start.Column; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                var propInfo = GetPropertyInfoByAttributeDescription(excelWorksheet.Cells[1, i].Value.ToString());
                // 如果属性类型是枚举类型, 就存储一份枚举成员与特性标记的映射键值对
                if (propInfo.PropertyType.IsEnum)
                {
                    enumProperties[propInfo.PropertyType.FullName] = GetEnumMemberDescriptionAndValueMapping(propInfo);
                }
                colIndexDictionary[i] = propInfo;
            }

            return colIndexDictionary;
        }

        /// <summary>
        /// 根据特性的描述查询属性
        /// </summary>
        /// <param name="description"></param>
        /// <returns></returns>
        private PropertyInfo GetPropertyInfoByAttributeDescription(string description)
        {
            return typeof(T).GetRuntimeProperties().FirstOrDefault(x => x.GetCustomAttribute<ExcelHeadDisplayAttribute>()?.HeadDisplay == description);
        }

        /// <summary>
        /// 获取枚举的特性标注与枚举值的键值对
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        private Dictionary<string, object> GetEnumMemberDescriptionAndValueMapping(PropertyInfo propertyInfo)
        {
            if (!propertyInfo.PropertyType.IsEnum) return null;
            var enumValues = propertyInfo.PropertyType.GetEnumValues();
            var enumDictionnary = new Dictionary<string, object>();
            foreach (var value in enumValues)
            {
                var memberAttr = propertyInfo.PropertyType.GetMember(value.ToString()).First().GetCustomAttribute<ExcelOptionItemDisplayAttribute>();
                enumDictionnary[memberAttr.OptionDisplay] = value;
            }

            return enumDictionnary;
        }

        /// <summary>
        /// 将把一个基础类型填充到对象属性中
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="propOwner"></param>
        /// <param name="value"></param>
        private void FillValueToProperty(ref PropertyInfo prop, object propOwner, object value, Dictionary<string, Dictionary<string, object>> enumDictionary)
        {
            if (prop.PropertyType == typeof(int))
            {
                prop.SetValue(propOwner, Convert.ToInt32(value));
                return;
            }

            if (prop.PropertyType == typeof(long))
            {
                prop.SetValue(propOwner, Convert.ToInt64(value));
                return;
            }

            if (prop.PropertyType == typeof(float))
            {
                prop.SetValue(propOwner, Convert.ToDouble(value));
                return;
            }

            if (prop.PropertyType == typeof(decimal))
            {
                prop.SetValue(propOwner, Convert.ToDecimal(value));
                return;
            }

            if (prop.PropertyType == typeof(double))
            {
                prop.SetValue(propOwner, Convert.ToDouble(value));
                return;
            }

            if (prop.PropertyType == typeof(bool))
            {
                prop.SetValue(propOwner, Convert.ToBoolean(value));
                return;
            }

            if (prop.PropertyType == typeof(char))
            {
                prop.SetValue(propOwner, Convert.ToChar(value));
                return;
            }

            if (prop.PropertyType == typeof(string))
            {
                prop.SetValue(propOwner, value.ToString());
                return;
            }

            if (prop.PropertyType.IsEnum)
            {
                prop.SetValue(propOwner, enumDictionary[prop.PropertyType.FullName][value.ToString()]);
            }
        }
    }
}
