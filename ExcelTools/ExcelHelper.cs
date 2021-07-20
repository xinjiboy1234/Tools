using ExcelTools.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelTools
{
    /// <summary>
    /// Excel 帮助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
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
            var colIndexDictionary =
                GetColumnIndexAndPropertiesMapping(workSheet, ref dictionaryOfEnumMemberAndDescription);

            var list = new List<T>();
            for (var i = 2; i <= workSheet.Dimension.End.Row; i++)
            {
                //var t = new T();
                var t = (T) Activator.CreateInstance(typeof(T));
                var props = t?.GetType().GetProperties();
                foreach (var (key, value) in colIndexDictionary)
                {
                    var prop = props?.FirstOrDefault(x => x.Name == value.Name);
                    FillValueToProperty(ref prop, t, workSheet.Cells[i, key].Value?.ToString(),
                        dictionaryOfEnumMemberAndDescription);
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
            var colIndexDictionary =
                GetColumnIndexAndPropertiesMapping(workSheet, ref dictionaryOfEnumMemberAndDescription);

            var list = new List<T>();
            for (var i = 2; i <= workSheet.Dimension.End.Row; i++)
            {
                //var t = new T();
                var t = (T) Activator.CreateInstance(typeof(T));
                var props = t?.GetType().GetProperties();
                foreach (var (key, value) in colIndexDictionary)
                {
                    var prop = props?.FirstOrDefault(x => x.Name == value.Name);
                    FillValueToProperty(ref prop, t, workSheet.Cells[i, key].Value.ToString(),
                        dictionaryOfEnumMemberAndDescription);
                }

                list.Add(t);
            }

            return list;
        }

        /// <summary>
        /// 保存Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <exception cref="Exception"></exception>
        public void SaveExcelFromCollection(List<T> data, string path)
        {
            if (data == null) throw new Exception("the data is required");

            using var excelPackage = new ExcelPackage(new FileInfo(path));
            var excelWorkSheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
            var dt = data.FirstOrDefault();
            // 先获取字段显示顺序
            var sortedPropertyInfos = SortingObjectProperties(dt).ToList();
            // 数据中包含的枚举类型的显示值与值的键值对，用于使用枚举值获取，字段显示值
            var enumDictionary = GetEumDictionary(dt);
            // 根据排序的字段顺序显示头部
            SetExcelHead(sortedPropertyInfos, ref excelWorkSheet);
            // 根据排序的字段顺序填充数据
            FillDataToExcel(data, ref excelWorkSheet, enumDictionary, sortedPropertyInfos);
            // 保存文件
            excelPackage.Save();
        }

        /// <summary>
        /// 设置头部
        /// </summary>
        /// <param name="sortedPropertyInfos"></param>
        /// <param name="excelWorkSheet"></param>
        private void SetExcelHead(List<PropertyInfo> sortedPropertyInfos, ref ExcelWorksheet excelWorkSheet)
        {
            var col = 1;
            foreach (var propertyInfo in sortedPropertyInfos)
            {
                excelWorkSheet.Cells[1, col].Value = propertyInfo
                    .GetCustomAttribute<ExcelHeadDisplayAttribute>()?
                    .HeadDisplay;

                col++;
            }
        }

        /// <summary>
        /// 将数据填充到Excel中
        /// </summary>
        /// <param name="data"></param>
        /// <param name="excelWorkSheet"></param>
        /// <param name="enumDictionary"></param>
        /// <param name="sortedPropertyInfos"></param>
        private void FillDataToExcel(List<T> data, ref ExcelWorksheet excelWorkSheet,
            Dictionary<string, Dictionary<string, object>> enumDictionary, List<PropertyInfo> sortedPropertyInfos)
        {
            var row = 2;
            foreach (var item in data)
            {
                var col = 1;
                foreach (var propertyInfo in sortedPropertyInfos)
                {
                    if (propertyInfo.PropertyType.IsEnum)
                    {
                        var tempDictionary = enumDictionary[propertyInfo.PropertyType.FullName ?? string.Empty];
                        SetOptionForExcelCell(ref excelWorkSheet, row, col, tempDictionary);
                        excelWorkSheet.Cells[row, col].Value = tempDictionary
                            .FirstOrDefault(x => x.Value.Equals(propertyInfo.GetValue(item))).Key;
                    }
                    else
                    {
                        excelWorkSheet.Cells[row, col].Value = propertyInfo.GetValue(item);
                    }

                    col++;
                }

                row++;
            }
        }

        /// <summary>
        /// excel 设置下拉框
        /// </summary>
        /// <param name="excelWorkSheet"></param>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <param name="enumMemberDictionary">枚举值与标注特性的键值对</param>
        /// <param name="prompt">提示信息</param>
        /// <param name="showPrompt">是否显示提示信息</param>
        private void SetOptionForExcelCell(ref ExcelWorksheet excelWorkSheet, int row, int col,
            Dictionary<string, object> enumMemberDictionary, string prompt = "",
            bool showPrompt = false)
        {
            var excelDataValidationList =
                excelWorkSheet.DataValidations.AddListValidation(excelWorkSheet.Cells[row, col].Address);

            foreach (var key in enumMemberDictionary.Keys)
            {
                excelDataValidationList.Formula.Values.Add(key);
            }

            excelDataValidationList.Prompt = prompt;
            excelDataValidationList.ShowInputMessage = showPrompt;
        }

        /// <summary>
        /// 获取枚举值与标注说明的键值对
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private Dictionary<string, Dictionary<string, object>> GetEumDictionary(T obj)
        {
            var result = new Dictionary<string, Dictionary<string, object>>();
            var propertyInfos = obj.GetType().GetProperties();
            foreach (var propertyInfo in propertyInfos)
            {
                if (!propertyInfo.PropertyType.IsEnum) continue;
                if (result.ContainsKey(propertyInfo.PropertyType.FullName ?? string.Empty)) continue;
                result.Add(propertyInfo.PropertyType.FullName ?? string.Empty,
                    GetEnumMemberDescriptionAndValueMapping(propertyInfo));
            }

            return result;
        }

        /// <summary>
        /// 获取excel表头顺序和与其对应类的属性键值对
        /// </summary>
        /// <param name="excelWorksheet"></param>
        /// <param name="enumProperties"></param>
        /// <returns></returns>
        private Dictionary<int, PropertyInfo> GetColumnIndexAndPropertiesMapping(ExcelWorksheet excelWorksheet,
            ref Dictionary<string, Dictionary<string, object>> enumProperties)
        {
            var colIndexDictionary = new Dictionary<int, PropertyInfo>();
            for (var i = excelWorksheet.Dimension.Start.Column; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                var propInfo = GetPropertyInfoByAttributeDescription(excelWorksheet.Cells[1, i].Value.ToString());
                // 如果属性类型是枚举类型, 就存储一份枚举成员与特性标记的映射键值对
                if (propInfo.PropertyType.IsEnum)
                {
                    if (string.IsNullOrEmpty(propInfo.PropertyType.FullName))
                        throw new Exception("property fullname is null or empty");
                    if (!enumProperties.ContainsKey(propInfo.PropertyType.FullName))
                    {
                        enumProperties.Add(propInfo.PropertyType.FullName,
                            GetEnumMemberDescriptionAndValueMapping(propInfo));
                    }
                }

                colIndexDictionary[i] = propInfo;
            }

            return colIndexDictionary;
        }

        /// <summary>
        /// 根据实体类上标注的字段显示顺序进行排序
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private IEnumerable<PropertyInfo> SortingObjectProperties(T obj)
        {
            return obj.GetType()
                .GetProperties()
                .OrderBy(x => x.GetCustomAttribute<ExcelHeadDisplayAttribute>()?.Index);
        }


        /// <summary>
        /// 根据特性的描述查询属性
        /// </summary>
        /// <param name="description"></param>
        /// <returns></returns>
        private PropertyInfo GetPropertyInfoByAttributeDescription(string description)
        {
            return typeof(T).GetRuntimeProperties().FirstOrDefault(x =>
                x.GetCustomAttribute<ExcelHeadDisplayAttribute>()?.HeadDisplay == description);
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
            var enumDictionary = new Dictionary<string, object>();
            foreach (var value in enumValues)
            {
                if (string.IsNullOrEmpty(value.ToString()))
                    throw new Exception($"the {propertyInfo.PropertyType}'s enumeration value is null");
                var memberAttr = propertyInfo.PropertyType.GetMember(value.ToString() ?? string.Empty).First()
                    .GetCustomAttribute<ExcelOptionItemDisplayAttribute>();
                if (memberAttr == null) throw new Exception($"the {propertyInfo.PropertyType}'s attribute is null");
                enumDictionary[memberAttr.OptionDisplay] = value;
            }

            return enumDictionary;
        }

        /// <summary>
        /// 将把一个基础类型填充到对象属性中
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="propOwner"></param>
        /// <param name="value"></param>
        /// <param name="enumDictionary"></param>
        private void FillValueToProperty(ref PropertyInfo prop, object propOwner, string value,
            Dictionary<string, Dictionary<string, object>> enumDictionary)
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
                prop.SetValue(propOwner, value);
                return;
            }

            if (!prop.PropertyType.IsEnum) return;
            if (string.IsNullOrEmpty(prop.PropertyType.FullName)) return;
            prop.SetValue(propOwner, enumDictionary[prop.PropertyType.FullName][value]);
        }
    }
}