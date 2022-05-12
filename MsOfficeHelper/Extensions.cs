using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

namespace OpenXmlBuilder
{
    public static class Extensions
    {
        /// <summary> IEnumerable to DataTable </summary>
        public static DataTable ToDataTable<T>(this IEnumerable<T> data)
        {
            var type = data.FirstOrDefault().GetType();
            List<PropertyInfo> properties =
               //typeof(T)
               type.GetProperties()
               .Where(pi => pi.GetCustomAttributes(typeof(XmlIgnoreAttribute), true).Length == 0)
               .ToList();

            DataTable table = new DataTable();
            foreach (PropertyInfo prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyInfo prop in properties)
                    row[prop.Name] = prop.GetValue(item, null) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }
    }
}
