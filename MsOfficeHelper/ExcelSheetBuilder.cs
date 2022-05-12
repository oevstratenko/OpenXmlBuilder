using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Xml.Serialization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlBuilder
{
    /// <summary> Builder for Excel Document, using OpenXml </summary>
    public sealed class ExcelSheetBuilder
    {
        #region private props and constructor

        /// <summary> sheet from Excel file </summary>
        private readonly SheetData sheet;
        /// <summary> merge cells setting </summary>
        private readonly MergeCells mergeCells;
        /// <summary> custom style setting </summary>
        private readonly List<CellPolygon> cellPolygon;
        /// <summary> cells with image </summary>
        private readonly List<ExcelHelper.ExportItem> cellImages;
        /// <summary> current row index </summary>
        private int rowIndex { get; set; }
        
        public ExcelSheetBuilder()
        {
            sheet       = new SheetData();
            mergeCells  = new MergeCells();
            cellPolygon = new List<CellPolygon>();
            cellImages  = new List<ExcelHelper.ExportItem>();

            rowIndex = 1;
        }

        #endregion

        #region public interface

        /// <summary> append empty line to sheet </summary>
        public ExcelSheetBuilder AppendLine()
        {
            rowIndex++;
            return this;
        }

        /// <summary> appent text to sheet </summary>
        public ExcelSheetBuilder AppendText(string aText, CellPolygon.CellStyle aStyle = null)
        {
            var idx = rowIndex++;

            var row = new Row {RowIndex = (uint)idx };
            if (aStyle?.Height != null)
            {
                row.Height = aStyle.Height;
                row.CustomHeight = true;
            }

            var cell = ExcelHelper.getCellTyped(aText);
            cell.CellReference = $"{ExcelHelper.getColumnA1Reference(1)}{idx}";

            row.AppendChild(cell);
            sheet.AppendChild(row);

            if (aStyle?.MergeCnt.GetValueOrDefault(0) > 0)
            {
                var mergeRangeReference = $"A{row.RowIndex}:{ExcelHelper.getColumnA1Reference(aStyle.MergeCnt.GetValueOrDefault(0))}{row.RowIndex}";
                mergeCells.Append(new MergeCell() { Reference = new StringValue(mergeRangeReference) });
            }

            if (aStyle != null)
            {
                cellPolygon.Add(
                    new CellPolygon(
                        new KeyValuePair<int, int>(idx, 1),
                        new KeyValuePair<int, int>(idx, 1),
                        aStyle));
            }

            return this;
        }

        /// <summary> Append table to sheet </summary>
        public ExcelSheetBuilder AppendTable<T>(IEnumerable<T> aData, TableSetting aTableSetting)
        {
            return this.AppendTable(aData, aTableSetting, x => new {});
        }

        /// <summary> Append table to sheet </summary>
        public ExcelSheetBuilder AppendTable<T, TResult>(IEnumerable<T> aData, TableSetting aTableSetting, Expression<Func<T, TResult>> aFields)
        {
            var _rowIndex = rowIndex;

            var type = typeof(T);

            //foramat cells for merge several cell as one by user settings
            Func<List<object>, int, List<object>> mergeApply = (f, cnt) =>
            {
                var newhFields = new List<object>();
                var i = 0;
                foreach (var val in f)
                {
                    var _mergeCnt = cnt - 1;
                    newhFields.Add(val);
                    i++;
                    while (_mergeCnt-- > 0)
                    {
                        i++;
                        newhFields.Add(string.Empty);
                    }

                    var mergeRangeReference = $"{ExcelHelper.getColumnA1Reference(i - cnt + 1)}{_rowIndex}:{ExcelHelper.getColumnA1Reference(i)}{_rowIndex}";
                    mergeCells.Append(new MergeCell {Reference = new StringValue(mergeRangeReference)});
                }

                return newhFields;
            };

            if (!(aFields.Body is NewExpression exp))
            {
                throw new ArgumentException("path wrong Expression syntax");
            }

            var properties = type.GetProperties()
                .Where(p =>
                    (exp.Members == null || exp.Members.Any(m => m.Name == p.Name))
                    && p.GetCustomAttributes(false).All(x => x.GetType() != typeof(XmlIgnoreAttribute))
                    )
                .Where(x => aTableSetting.ExportedColumns == null || aTableSetting.ExportedColumns.Contains(x.Name))
                .OrderBy(x => aTableSetting.ExportedColumns == null ? null : (int?)Array.IndexOf(aTableSetting.ExportedColumns, x.Name));

            var inc = 0;

            var hFields = properties.Select(p => (p.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as DisplayNameAttribute)?.DisplayName ?? p.Name).ToList<object>();
            if (!aTableSetting.HideHeader)
            {
                if (aTableSetting.HeaderCellStyle.MergeCnt.GetValueOrDefault(0) > 0)
                {
                    hFields = mergeApply(hFields, aTableSetting.HeaderCellStyle.MergeCnt.GetValueOrDefault(0));
                }

                var row = ExcelHelper.getDataRow(hFields, rowIndex++);
                
                if (aTableSetting.HeaderCellStyle.Height != null)
                {
                    row.Height = aTableSetting.HeaderCellStyle.Height;
                    row.CustomHeight = true;
                }
                sheet.AppendChild(row);

                //header
                cellPolygon.Add(
                    new CellPolygon(
                        new KeyValuePair<int, int>(_rowIndex + inc, 1),
                        new KeyValuePair<int, int>(_rowIndex + inc, hFields.Count()),
                        aTableSetting.HeaderCellStyle));

                inc++;
            }

            //data
            cellPolygon.Add(
                new CellPolygon(
                    new KeyValuePair<int, int>(_rowIndex + inc, 1),
                    new KeyValuePair<int, int>(_rowIndex + aData.Count(), hFields.Count() * aTableSetting.DataCellStyle.MergeCnt.GetValueOrDefault(1)),
                    aTableSetting.DataCellStyle));

            //foreach (var values in aData.Select(item => properties.Select(p => p.GetValue(item, null)).ToList()).ToList())
            foreach (var item in aData.Select(item =>
            {
                var i = 0;
                return properties.Select(p => new ExcelHelper.ExportItem
                {
                    value = p.GetValue(item, null),
                    isImage = p.GetCustomAttributes(typeof(ExcelHelper.ExcelImagePathAttribute), false).Any(),
                    imageWidthPx = (p.GetCustomAttributes(typeof(ExcelHelper.ExcelImagePathAttribute), true).FirstOrDefault() as ExcelHelper.ExcelImagePathAttribute)?.WidthPx,
                    colNo = ++i
                }).ToList();
            }).ToList())
            {
                foreach (var _item in item.Where(x => x.isImage).ToList())
                {
                    if (string.IsNullOrEmpty(_item.value?.ToString()))
                    {
                        continue;
                    }
                    //_item.value = string.Empty;
                    _item.rowNo = rowIndex;
                    cellImages.Add(_item);
                }

                var _values = item.Select(x => x.value).ToList();
                if (aTableSetting.DataCellStyle.MergeCnt.GetValueOrDefault(0) > 0)
                {
                    _rowIndex += inc;
                    _values = mergeApply(_values, aTableSetting.DataCellStyle.MergeCnt.GetValueOrDefault(0));
                }

                var row = ExcelHelper.getDataRow(_values, rowIndex++);
                if (aTableSetting.DataCellStyle.Height != null)
                {
                    row.Height = aTableSetting.DataCellStyle.Height;
                    row.CustomHeight = true;
                }
                sheet.AppendChild(row);
            }

            return this;
        }

        /// <summary> process build excel document </summary>
        public void SaveToFile(string aFileName, string aSheetName, bool aOverrideIfExists = true)
        {
            using (var sd = !aOverrideIfExists && File.Exists(aFileName)
                ? SpreadsheetDocument.Open(aFileName, true)
                : SpreadsheetDocument.Create(aFileName, SpreadsheetDocumentType.Workbook))
            {
                var worksheet = ExcelHelper.createWorksheet(sd, sheet, aSheetName);

                long ticks = DateTime.Now.Ticks;
                int i = 0;
                foreach (var item in cellPolygon)
                {
                    //#Debug.WriteLine($"{i++} : {(DateTime.Now.Ticks - ticks) / TimeSpan.TicksPerMillisecond}");
                    //#ticks = DateTime.Now.Ticks;

                    if (item.Style.BorderStyle != BorderStyleValues.None)
                    {
                        ExcelHelper.setBorder(sd.WorkbookPart, worksheet, item.CellFrom, item.CellTo, item.Style.BorderStyle);
                    }

                    if (item.Style.BackColorRgb != null)
                    {
                        ExcelHelper.setFill(sd.WorkbookPart, worksheet, item.CellFrom, item.CellTo, item.Style.BackColorRgb);
                    }

                    if (item.Style.Alignment != null)
                    {
                        ExcelHelper.setAlignment(sd.WorkbookPart, worksheet, item.CellFrom, item.CellTo, item.Style.Alignment);
                    }

                    if (item.Style.Font != null)
                    {
                        ExcelHelper.setFont(sd.WorkbookPart, worksheet, item.CellFrom, item.CellTo, item.Style.Font);
                    }
                }

                if (mergeCells.Any())
                {
                    worksheet.InsertAfter(mergeCells, sheet);
                }

                foreach (var item in cellImages)
                {
                    try
                    {
                        using (var imageStream = new FileStream(item.value?.ToString(), FileMode.Open, FileAccess.Read))
                        {
                            ExcelHelper.AddImage(worksheet.WorksheetPart, imageStream, item.value?.ToString(), item.colNo, item.rowNo, item.imageWidthPx);
                        }
                        ExcelHelper.setValue(worksheet, item.colNo, item.rowNo, string.Empty);
                    }
                    catch (FileNotFoundException e)
                    {
                        ExcelHelper.setValue(worksheet, item.colNo, item.rowNo, e.Message);
                    }
                }
                
                worksheet.Save();
            }
        }

        #endregion

        #region excel style prop

        /// <summary> excel table style </summary>
        public class TableSetting
        {
            public TableSetting()
            {
                HeaderCellStyle = new CellPolygon.CellStyle()
                {
                    BorderStyle = BorderStyleValues.Thin
                };
                DataCellStyle = new CellPolygon.CellStyle()
                {
                    //BorderStyle = BorderStyleValues.Thin//too much time for processing if big datatable
                    BorderStyle = BorderStyleValues.None
                };
            }
            /// <summary> table header style </summary>
            public CellPolygon.CellStyle HeaderCellStyle { get; set; }
            /// <summary> table body style </summary>
            public CellPolygon.CellStyle DataCellStyle { get; set; }
            /// <summary> need to hide header </summary>
            public bool HideHeader { get; set; }

            /// <summary> Columns for export </summary>
            public string[] ExportedColumns { get; set; } 
        }

        /// <summary> Cell formated area (example A1:D6) </summary>
        public class CellPolygon
        {
            public CellPolygon(KeyValuePair<int, int> aCellFrom, KeyValuePair<int, int> aCellTo, CellStyle aStyle)
            {
                CellFrom = aCellFrom;
                CellTo   = aCellTo;
                Style    = aStyle;
            }
            /// <summary> Cell address of Begin of cell Area (row num,column num) </summary>
            public KeyValuePair<int, int>   CellFrom{ get; private set; }
            /// <summary> Cell address of End of cell Area (row num,column num) </summary>
            public KeyValuePair<int, int>   CellTo  { get; private set; }

            public CellStyle                Style   { get; private set; }

            /// <summary> Cell foramated properties </summary>
            public class CellStyle
            {
                /// <summary> BorderStyle </summary>
                public BorderStyleValues BorderStyle    { get; set; }
                /// <summary> Height </summary>
                public int?              Height         { get; set; }
                /// <summary> Cells quantity for merge in one cell </summary>
                public int?              MergeCnt       { get; set; }
                /// <summary> Background color </summary>
                public string            BackColorRgb   { get; set; }
                /// <summary> Alignment </summary>
                public Alignment         Alignment      { get; set; }
                /// <summary> Font </summary>
                public Font              Font           { get; set; }
            }
        }

        #endregion
    }
}