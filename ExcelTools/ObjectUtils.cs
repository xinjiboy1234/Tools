using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelTools
{
    /// <summary>
    /// 对象工具类
    /// </summary>
    public static class ObjectUtils
    {
        /// <summary>
        /// 根据特性的描述查询属性
        /// </summary>
        /// <param name="description"></param>
        /// <returns></returns>
        public static PropertyInfo GetPropertyInfoByAttributeDescription<T>(string description) where T : class, new()
        {
            return typeof(T).GetProperties().FirstOrDefault(x => (x.GetCustomAttribute(typeof(HeadAttribute)) as HeadAttribute)?.Head == description);
        }

        /// <summary>
        /// 将把一个基础类型填充到对象属性中
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="propOwner"></param>
        /// <param name="value"></param>
        public static void FillValueToProperty(ref PropertyInfo prop, object propOwner, object value)
        {
            if(prop.PropertyType == typeof(int))
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
        }
    }
}
