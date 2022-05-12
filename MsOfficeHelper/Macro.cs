using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace OpenXmlBuilder
{
    /// <summary> Data provider for Macro implementation </summary>
    public interface IMacroDataProvider
    {
        IEnumerable<object> GetData(string aKey);
    }

    /// <summary> Get macro Value with parameters (NameWithParams) by key from datasource (IMacroDataProvider) </summary>
    public class Macro
    {
        /// <summary> Data provider </summary>
        private readonly IMacroDataProvider dataProvider;

        public Macro(IMacroDataProvider aDataProvider)
        {
            dataProvider = aDataProvider;
        }

        /// <summary> macro result type (table, scalar value) </summary>
        private struct resultType
        {
            /// <summary> table </summary>
            public const string table = ":table:";
            /// <summary> scalar value </summary>
            public const string value = ":value:";
        }

        public const string formatStart = "{{format:";
        public const string end = "}}";

        public const string captionStart = "{{caption:";
        //image path macro property name ({{image:<width>x<height>)
        public const string imageStart = "{{image:";
        //column width
        public const string widthStart = "{{width:";

        /// <summary> Get macro value </summary>
        public object GetValue(string aMacro)
        {
            if (dataProvider == null)
            {
                throw new ArgumentNullException("IMacroDataProvider not implement");
            }

            object result = null;

            var macroName = aMacro.Substring(0, aMacro.IndexOf(":", StringComparison.InvariantCultureIgnoreCase));

            var list = dataProvider.GetData(macroName);

            if (list == null || !list.Any())
            {
                return null;
            }

            var dataTable = list.ToDataTable();

            if (aMacro.Contains(resultType.value))
            {
                var resTypeIdx = aMacro.IndexOf(resultType.value, StringComparison.InvariantCultureIgnoreCase) + resultType.value.Length;
                var columnNameFormat = GetColNameAndParams(aMacro.Substring(resTypeIdx, aMacro.Length - resTypeIdx));
                result = SetFormat(dataTable.Rows[0][columnNameFormat.Name], columnNameFormat.Format);
            }
            if (aMacro.Contains(resultType.table))
            {
                var resTypeIdx = aMacro.IndexOf(resultType.table, StringComparison.InvariantCultureIgnoreCase) + resultType.value.Length;
                var columnNamesString = aMacro.Substring(resTypeIdx, aMacro.Length - resTypeIdx);
                var columnNameFormats = columnNamesString
                    .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(GetColNameAndParams)
                    .ToList();

                var resultTable = new DataTable();
                foreach (var item in columnNameFormats)
                {
                    var col         = new DataColumn(item.Name);
                    col.DataType    = typeof(string);
                    col.AllowDBNull = true;
                    col.Caption     = item.Caption;
                    col.ExtendedProperties.Add("width", item.Width);

                    resultTable.Columns.Add(col);
                }

                foreach (DataRow row in dataTable.Rows)
                {
                    var newRow = resultTable.Rows.Add();
                    foreach (DataColumn col in resultTable.Columns)
                    {
                        newRow[col.ColumnName] = SetFormat(row[col.ColumnName], columnNameFormats.FirstOrDefault(x => x.Name == col.ColumnName).Format);
                    }
                }

                var imageArray = columnNameFormats.Where(x => x.Image != null).Select(x =>
                {
                    int? width = null, heigth = null;

                    var format = columnNameFormats.FirstOrDefault(xx => xx.Name == x.Name)?.Image;
                    if (!string.IsNullOrWhiteSpace(format))
                    {
                        var arr = format?.Split(new [] {'x'}).Select(xx => Convert.ToInt32(xx)).ToArray();
                        width = arr[0];
                        heigth = arr[1];
                    }

                    return new WordHelper.ImageColumn {Name = x.Name, Height = heigth, Width = width};
                }).ToArray();
                resultTable.ExtendedProperties.Add("imageArray", imageArray);

                result = resultTable;
            }

            return result;
        }

        /// <summary> Get macro name and parameters (colName{{format:}}, etc) </summary>
        private NameWithParams GetColNameAndParams(string aValue)
        {
            string colName = aValue;

            string format = null;
            if (aValue.IndexOf(formatStart, StringComparison.InvariantCultureIgnoreCase) >= 0)
            {
                var formatStartIndex    = aValue.IndexOf(formatStart, StringComparison.InvariantCultureIgnoreCase);
                var formatEndIndex      = aValue.IndexOf(end, formatStartIndex, StringComparison.InvariantCultureIgnoreCase);

                format = aValue.Substring(formatStartIndex + formatStart.Length, formatEndIndex - formatStartIndex - formatStart.Length);
                colName = aValue?.Replace(formatStart + format + end, string.Empty);
            }

            string image = null;
            if (colName.IndexOf(imageStart, StringComparison.InvariantCultureIgnoreCase) >= 0)
            {
                var startIndex = colName.IndexOf(imageStart, StringComparison.InvariantCultureIgnoreCase);
                var endIndex = colName.IndexOf(end, startIndex, StringComparison.InvariantCultureIgnoreCase);

                image = colName.Substring(startIndex + imageStart.Length, endIndex - startIndex - imageStart.Length);
                colName = colName?.Replace(imageStart + image + end, string.Empty);
            }

            string caption = null;
            if (colName.IndexOf(captionStart, StringComparison.InvariantCultureIgnoreCase) >= 0)
            {
                var startIndex = colName.IndexOf(captionStart, StringComparison.InvariantCultureIgnoreCase);
                var endIndex = colName.IndexOf(end, startIndex, StringComparison.InvariantCultureIgnoreCase);

                caption = colName.Substring(startIndex + captionStart.Length, endIndex - startIndex - captionStart.Length);
                colName = colName?.Replace(captionStart + caption + end, string.Empty);
            }
            
            string width = null;
            if (colName.IndexOf(widthStart, StringComparison.InvariantCultureIgnoreCase) >= 0)
            {
                var startIndex = colName.IndexOf(widthStart, StringComparison.InvariantCultureIgnoreCase);
                var endIndex = colName.IndexOf(end, startIndex, StringComparison.InvariantCultureIgnoreCase);

                width = colName.Substring(startIndex + widthStart.Length, endIndex - startIndex - widthStart.Length);
                colName = colName?.Replace(widthStart + width + end, string.Empty);
            }

            return new NameWithParams(colName, format, caption, image, width);
        }

        /// <summary> Apply Value Format </summary>
        private object SetFormat(object aValue, string aFormat)
        {
            if (aValue == null || string.IsNullOrWhiteSpace(aFormat))
            {
                return aValue;
            }

            if (aValue is DateTime?)
            {
                return ((DateTime)aValue).ToString(aFormat);
            }
            if (aValue is decimal?)
            {
                return ((decimal)aValue).ToString(aFormat);
            }

            return aValue;
        }

        /// <summary> macro properties </summary>
        private class NameWithParams
        {
            public NameWithParams(string aName, string aFormat, string aCaption, string aImage, string aWidth)
            {
                this.Name       = System.Text.RegularExpressions.Regex.Replace(aName, @"\{{([^\}]+)\}}", "");//clear undifined macro attribute
                this.Format     = aFormat;
                this.Caption    = aCaption;
                this.Image      = aImage;
                this.Width      = aWidth;
            }

            /// <summary> Macro Name </summary>
            public string Name     { get; set; }
            /// <summary> Column display name (Table only) </summary>
            public string Caption  { get; set; }
            /// <summary> Macro Value Format </summary>
            public string Format   { get; set; }
            /// <summary> Image properties (width, height, etc) </summary>
            public string Image    { get; set; }
            /// <summary> Column width (Table only) </summary>
            public string Width    { get; set; }
        }
    }
}