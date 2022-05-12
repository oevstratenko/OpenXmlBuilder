using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Drawing;
using System.Drawing.Imaging;

namespace OpenXmlBuilder
{
    public static class ExcelHelper
    {
        /// <summary> Custom styles </summary>
        /// <remarks> https://www.codeproject.com/Articles/97307/Using-C-and-Open-XML-SDK-for-Microsoft-Office </remarks>
        private class CustomStylesheet : Stylesheet
        {
            /// <summary>
            /// ID  Format Code
            /// ID  Format Code
            /// 0   General
            /// 1   0
            /// 2   0.00
            /// 3   #,##0
            /// 4   #,##0.00
            /// 9   0%
            /// 10  0.00%
            /// 11  0.00E+00
            /// 12  # ?/?
            /// 13  # ??/??
            /// 14  d/m/yyyy
            /// 15  d-mmm-yy
            /// 16  d-mmm
            /// 17  mmm-yy
            /// 18  h:mm tt
            /// 19  h:mm:ss tt
            /// 20  H:mm
            /// 21  H:mm:ss
            /// 22  m/d/yyyy H:mm
            /// 37  #,##0 ;(#,##0)
            /// 38  #,##0 ;[Red](#,##0)
            /// 39  #,##0.00;(#,##0.00)
            /// 40  #,##0.00;[Red](#,##0.00)
            /// 45  mm:ss
            /// 46  [h]:mm:ss
            /// 47  mmss.0
            /// 48  ##0.0E+0
            /// 49  @
            /// </summary>
            public struct Formats
            {
                public const int General = 0;
                public const int Number = 1;
                public const int Decimal = 2;
                public const int Decimal2 = 3;
                public const int Decimal3 = 4;
                public const int Currency = 164;
                public const int Accounting2 = 43;
                public const int Accounting = 44;
                public const int DateShort = 14;
                public const int DateLong = 165;
                public const int Time = 166;
                public const int Percentage = 10;
                public const int Fraction = 12;
                public const int Scientific = 11;
                public const int Text = 49;

            }

            public CustomStylesheet()
            {
                var fonts = new Fonts();
                var font = new DocumentFormat.OpenXml.Spreadsheet.Font();
                var fontName = new FontName { Val = StringValue.FromString("Calibri") };
                var fontSize = new FontSize { Val = DoubleValue.FromDouble(10) };
                font.FontName = fontName;
                font.FontSize = fontSize;
                fonts.Append(font);
                //Font Index 1
                font = new DocumentFormat.OpenXml.Spreadsheet.Font();
                fontName = new FontName { Val = StringValue.FromString("Calibri") };
                fontSize = new FontSize { Val = DoubleValue.FromDouble(10) };
                font.FontName = fontName;
                font.FontSize = fontSize;
                font.Bold = new Bold();
                fonts.Append(font);
                fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
                var fills = new Fills();
                var fill = new Fill();
                var patternFill = new PatternFill { PatternType = PatternValues.None };
                fill.PatternFill = patternFill;
                fills.Append(fill);
                fill = new Fill();
                patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
                fill.PatternFill = patternFill;
                fills.Append(fill);
                //Fill index  2
                fill = new Fill();
                patternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor()
                };
                patternFill.ForegroundColor =
                   TranslateForeground(System.Drawing.Color.LightBlue);
                patternFill.BackgroundColor =
                    new BackgroundColor { Rgb = patternFill.ForegroundColor.Rgb };
                fill.PatternFill = patternFill;
                fills.Append(fill);
                //Fill index  3
                fill = new Fill();
                patternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor()
                };
                patternFill.ForegroundColor =
                   TranslateForeground(System.Drawing.Color.DodgerBlue);
                patternFill.BackgroundColor =
                   new BackgroundColor { Rgb = patternFill.ForegroundColor.Rgb };
                fill.PatternFill = patternFill;
                fills.Append(fill);
                fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
                var borders = new Borders();
                var border = new Border
                {
                    LeftBorder = new LeftBorder(),
                    RightBorder = new RightBorder(),
                    TopBorder = new TopBorder(),
                    BottomBorder = new BottomBorder(),
                    DiagonalBorder = new DiagonalBorder()
                };
                borders.Append(border);
                //All Boarder Index 1
                border = new Border
                {
                    LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                    RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                    TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                    BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                    DiagonalBorder = new DiagonalBorder()
                };
                borders.Append(border);
                //Top and Bottom Boarder Index 2
                border = new Border
                {
                    LeftBorder = new LeftBorder(),
                    RightBorder = new RightBorder(),
                    TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                    BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                    DiagonalBorder = new DiagonalBorder()
                };
                borders.Append(border);
                borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
                var cellStyleFormats = new CellStyleFormats();
                var cellFormat = new CellFormat
                {
                    NumberFormatId = 0,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0
                };
                cellStyleFormats.Append(cellFormat);
                cellStyleFormats.Count =
                   UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
                uint iExcelIndex = 164;
                var numberingFormats = new NumberingFormats();
                var cellFormats = new CellFormats();
                cellFormat = new CellFormat
                {
                    NumberFormatId = 0,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0
                };
                cellFormats.Append(cellFormat);
                var nformatDateTime = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                    FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
                };
                numberingFormats.Append(nformatDateTime);
                var nformat4Decimal = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                    FormatCode = StringValue.FromString("#,##0.0000")
                };
                numberingFormats.Append(nformat4Decimal);
                var nformat2Decimal = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                    FormatCode = StringValue.FromString("#,##0.00")
                };
                numberingFormats.Append(nformat2Decimal);
                var nformatForcedText = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(iExcelIndex),
                    FormatCode = StringValue.FromString("@")
                };
                numberingFormats.Append(nformatForcedText);
                // index 1
                // Cell Standard Date format 
                cellFormat = new CellFormat
                {
                    NumberFormatId = 14,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 2
                // Cell Standard Number format with 2 decimal placing
                cellFormat = new CellFormat
                {
                    NumberFormatId = 4,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 3
                // Cell Date time custom format
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformatDateTime.NumberFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 4
                // Cell 4 decimal custom format
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformat4Decimal.NumberFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 5
                // Cell 2 decimal custom format
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformat2Decimal.NumberFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 6
                // Cell forced number text custom format
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformatForcedText.NumberFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 7
                // Cell text with font 12 
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformatForcedText.NumberFormatId,
                    FontId = 1,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 8
                // Cell text
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformatForcedText.NumberFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 1,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 9
                // Coloured 2 decimal cell text
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformat2Decimal.NumberFormatId,
                    FontId = 0,
                    FillId = 2,
                    BorderId = 2,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 10
                // Coloured cell text
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformatForcedText.NumberFormatId,
                    FontId = 0,
                    FillId = 2,
                    BorderId = 2,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                // Index 11
                // Coloured cell text
                cellFormat = new CellFormat
                {
                    NumberFormatId = nformatForcedText.NumberFormatId,
                    FontId = 1,
                    FillId = 3,
                    BorderId = 2,
                    FormatId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                };
                cellFormats.Append(cellFormat);
                numberingFormats.Count =
                  UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
                cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
                this.Append(numberingFormats);
                this.Append(fonts);
                this.Append(fills);
                this.Append(borders);
                this.Append(cellStyleFormats);
                this.Append(cellFormats);
                var css = new CellStyles();
                var cs = new CellStyle
                {
                    Name = StringValue.FromString("Normal"),
                    FormatId = 0,
                    BuiltinId = 0
                };
                css.Append(cs);
                css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
                this.Append(css);
                var dfs = new DifferentialFormats { Count = 0 };
                this.Append(dfs);
                var tss = new TableStyles
                {
                    Count = 0,
                    DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                    DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
                };
                this.Append(tss);
            }
            private static ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
            {
                return new ForegroundColor()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value =
                                  System.Drawing.ColorTranslator.ToHtml(
                                  System.Drawing.Color.FromArgb(
                                      fillColor.A,
                                      fillColor.R,
                                      fillColor.G,
                                      fillColor.B)).Replace("#", "")
                    }
                };
            }
        }

        internal static Worksheet createWorksheet(SpreadsheetDocument aDocument, SheetData aSheetData, string aTableName)
        {
            if (aDocument.WorkbookPart == null)
            {
                var workbookPart = aDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Create Styles and Insert into Workbook
                var stylesPart = workbookPart.Workbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);

                workbookPart.Workbook.Sheets = new Sheets();
            }

            WorksheetPart wsp = aDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(aSheetData);

            //wsp.Worksheet.AppendChild(aSheetData);

            wsp.Worksheet.Save();

            UInt32 sheetId;

            // If this is the first sheet, the ID will be 1. If this is not the first sheet, we calculate the ID based on the number of existing
            // sheets + 1.
            if (aDocument.WorkbookPart.Workbook.Sheets == null)
            {
                aDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                sheetId = 1;
            }
            else
            {
                sheetId = Convert.ToUInt32(aDocument.WorkbookPart.Workbook.Sheets.Count() + 1);
            }

            // Create the new sheet and add it to the workbookpart
            aDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = aDocument.WorkbookPart.GetIdOfPart(wsp),
                    SheetId = sheetId,
                    Name = aTableName
                }
            );

            //_cols = new Columns(); // Created to allow bespoke width columns
            // Save our changes
            aDocument.WorkbookPart.Workbook.Save();

            return wsp.Worksheet;// wsp;
        }

        internal static Row getDataRow(IEnumerable<object> aData, int rowIndex)
        {
            var result = new Row();
            result.RowIndex = (uint)rowIndex;
            var colNo = 1;
            foreach (var value in aData)
            {
                var cell = getCellTyped(value);
                cell.CellReference = $"{getColumnA1Reference(colNo++)}{rowIndex}";
                result.AppendChild(cell);
            }
            return result;
        }
        
        private static Row getDataRow(DataTable aData, int rowIndex)
        {
            var result = new Row();
            foreach (var col in aData.Columns)
            {
                result.AppendChild(
                    getCellTyped(
                        aData.Rows[rowIndex][col.ToString()]
                        )
                    );
            }
            return result;
        }

        internal static Cell getCellTyped(object aVal)
        {
            var cell = new Cell();

            if (aVal is int || aVal is long)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToString(aVal));
            }
            else if (aVal is decimal)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToString(aVal)?.Replace(",", "."));
                cell.StyleIndex = CustomStylesheet.Formats.Decimal;
            }
            else
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(Convert.ToString(aVal));
            }

            return cell;
        }

        internal static string getColumnA1Reference(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        #region append cell style

        private static CellFormat getCellFormat(WorkbookPart workbookPart, uint styleIndex)
        {
            return workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().First().Elements<CellFormat>().ElementAt((int)styleIndex);
        }

        private static Cell getCell(Worksheet workSheet, string cellAddress)
        {
            var cells = workSheet.Descendants<Cell>();

            var res = cells.SingleOrDefault(c => cellAddress.Equals(c.CellReference));

            return res;
        }

        private static uint insertCellFormat(WorkbookPart workbookPart, CellFormat cellFormat)
        {
            CellFormats cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().First();
            cellFormats.Append(cellFormat);
            return cellFormats.Count++;
        }

        #region setBorder

        private static Border generateBorder(BorderStyleValues aBorderStyle)
        {
            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = aBorderStyle };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color1);

            RightBorder rightBorder2 = new RightBorder() { Style = aBorderStyle };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color2);

            TopBorder topBorder2 = new TopBorder() { Style = aBorderStyle };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color3);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = aBorderStyle };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color4);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            return border2;
        }
        
        #endregion

        #region setFill

        private static Fill generateFill(string aBackColorRgb)
        {
            Fill fill = new Fill();

            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = aBackColorRgb };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill.Append(foregroundColor1);
            patternFill.Append(backgroundColor1);

            fill.Append(patternFill);

            return fill;
        }

        private static uint insertFill(WorkbookPart workbookPart, Fill fill)
        {
            Fills fills = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Fills>().First();
            fills.Append(fill);
            return (uint)fills.Count++;
        }

        public static void setFill(WorkbookPart workbookPart, Worksheet workSheet, string cellAddress, string aColorRgb)
        {
            Cell cell = getCell(workSheet, cellAddress);
            if (cell == null)
            {
                return;
            }

            CellFormat cellFormat = cell.StyleIndex != null ? getCellFormat(workbookPart, cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            cellFormat.FillId = insertFill(workbookPart, generateFill(aColorRgb));

            cell.StyleIndex = insertCellFormat(workbookPart, cellFormat);
        }

        public static void setValue(Worksheet workSheet, int col, int row, string aValue)
        {
            var address = $"{getColumnA1Reference(col)}{row}";
            var cell = getCell(workSheet, address);
            cell.CellValue = new CellValue(aValue);
        }

        #endregion

        #region setAlignment

        public static void setAlignment(WorkbookPart workbookPart, Worksheet workSheet, string cellAddress, Alignment aAlignment)
        {
            Cell cell = getCell(workSheet, cellAddress);
            if (cell == null)
            {
                return;
            }

            CellFormat cellFormat = cell.StyleIndex != null ? getCellFormat(workbookPart, cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            cellFormat.Alignment = aAlignment.CloneNode(true) as Alignment;

            cell.StyleIndex = insertCellFormat(workbookPart, cellFormat);
        }

        #endregion

        #region setFont

        private static uint insertFont(WorkbookPart workbookPart, Font fill)
        {
            Fonts fonts = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().First();
            fonts.Append(fill);
            return (uint)fonts.Count++;
        }
        
        public static void setFont(WorkbookPart workbookPart, Worksheet workSheet, string cellAddress, Font aFont)
        {
            Cell cell = getCell(workSheet, cellAddress);
            if (cell == null)
            {
                return;
            }

            CellFormat cellFormat = cell.StyleIndex != null ? getCellFormat(workbookPart, cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            var font = aFont.CloneNode(true) as Font;
            cellFormat.FontId = insertFont(workbookPart, font);

            cell.StyleIndex = insertCellFormat(workbookPart, cellFormat);
        }

        #endregion

        #region setAlignment,setFill,setBorder,setFont for cells Area (example A1:C12)

        public static void setAlignment(WorkbookPart workbookPart, Worksheet workSheet, KeyValuePair<int, int>? aCellFrom, KeyValuePair<int, int>? aCellTo, Alignment aAlignment)
        {
            aCellTo = aCellTo ?? aCellFrom;
            for (var row = aCellFrom.Value.Key; row <= aCellTo.Value.Key; row++)
            {
                for (var col = aCellFrom.Value.Value; col <= aCellTo.Value.Value; col++)
                {
                    var address = $"{getColumnA1Reference(col)}{row}";
                    setAlignment(workbookPart, workSheet, address, aAlignment);
                }
            }
        }

        public static void setFill(WorkbookPart workbookPart, Worksheet workSheet, KeyValuePair<int, int>? aCellFrom, KeyValuePair<int, int>? aCellTo, string aColorRgb)
        {
            aCellTo = aCellTo ?? aCellFrom;
            for (var row = aCellFrom.Value.Key; row <= aCellTo.Value.Key; row++)
            {
                for (var col = aCellFrom.Value.Value; col <= aCellTo.Value.Value; col++)
                {
                    var address = $"{getColumnA1Reference(col)}{row}";
                    setFill(workbookPart, workSheet, address, aColorRgb?.Replace("#", string.Empty));
                }
            }
        }

        public static void setBorder(WorkbookPart workbookPart, Worksheet workSheet, KeyValuePair<int, int>? aCellFrom, KeyValuePair<int, int>? aCellTo, BorderStyleValues aBorderStyle)
        {
            uint? styleIdx = null;

            aCellTo = aCellTo ?? aCellFrom;
            for (var row = aCellFrom.Value.Key; row <= aCellTo.Value.Key; row++)
            {
                for (var col = aCellFrom.Value.Value; col <= aCellTo.Value.Value; col++)
                {
                    var address = $"{getColumnA1Reference(col)}{row}";

                    Cell cell = getCell(workSheet, address);
                    if (cell == null)
                    {
                        continue;
                    }

                    if (styleIdx == null)
                    {
                        styleIdx = getBorder(workbookPart, workSheet, cell, aBorderStyle);
                    }

                    cell.StyleIndex = styleIdx;
                }
            }
        }

        public static uint getBorder(WorkbookPart workbookPart, Worksheet workSheet, Cell cell, BorderStyleValues aBorderStyle = BorderStyleValues.Thin)
        {
            CellFormat cellFormat = cell.StyleIndex != null ? getCellFormat(workbookPart, cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            var border = generateBorder(aBorderStyle);

            Borders borders = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().First();
            borders.Append(border);
            cellFormat.BorderId = borders.Count++;

            return insertCellFormat(workbookPart, cellFormat);
        }

        public static void setFont(WorkbookPart workbookPart, Worksheet workSheet, KeyValuePair<int, int>? aCellFrom, KeyValuePair<int, int>? aCellTo, Font aFont)
        {
            aCellTo = aCellTo ?? aCellFrom;
            for (var row = aCellFrom.Value.Key; row <= aCellTo.Value.Key; row++)
            {
                for (var col = aCellFrom.Value.Value; col <= aCellTo.Value.Value; col++)
                {
                    var address = $"{getColumnA1Reference(col)}{row}";
                    setFont(workbookPart, workSheet, address, aFont);
                }
            }
        }

        #endregion

        #endregion

        /// <summary> Export IEnumerable to Excel (PropertyName = ColumnName) </summary>
        public static void exportToExcel<T>(this IEnumerable<T> aEnumerable, string aFileName, string aTableName, bool aImportWithXmlJsonIgnore = false, bool aOverrideIfExists = true) where T : class
        {
            using (var spreadsheetDocument =
                !aOverrideIfExists && File.Exists(aFileName)
                ? SpreadsheetDocument.Open(aFileName, true)
                : SpreadsheetDocument.Create(aFileName, SpreadsheetDocumentType.Workbook))
            {
                var sheetData = new SheetData();
                var worksheet = createWorksheet(spreadsheetDocument, sheetData, aTableName);

                var type = typeof(T);
                var properties = type.GetProperties().Where(p => aImportWithXmlJsonIgnore || p.GetCustomAttributes(false).All(x => x.GetType() != typeof(XmlIgnoreAttribute)));
                
                int rowIndex = 1;
                sheetData.AppendChild(getDataRow(properties.Select(p => (p.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as DisplayNameAttribute)?.DisplayName ?? p.Name), rowIndex++));

                foreach (var item in aEnumerable.Select(item =>
                {
                    var i = 0;
                    return properties.Select(p => new ExportItem
                    {
                        value       = p.GetValue(item, null),
                        isImage     = p.GetCustomAttributes(typeof(ExcelImagePathAttribute), false).Any(),
                        imageWidthPx= (p.GetCustomAttributes(typeof(ExcelImagePathAttribute), true).FirstOrDefault() as ExcelImagePathAttribute)?.WidthPx,
                        colNo = ++i
                    }).ToList();
                }).ToList())
                {
                    foreach (var _item in item.Where(x => x.isImage))
                    {
                        if (string.IsNullOrEmpty(_item.value?.ToString()))
                        {
                            continue;
                        }
                        try
                        {
                            using (var imageStream = new FileStream(_item.value?.ToString(), FileMode.Open, FileAccess.Read))
                            {
                                AddImage(worksheet.WorksheetPart, imageStream, _item.value?.ToString(), _item.colNo, rowIndex, _item.imageWidthPx);

                                _item.value = string.Empty;
                            }
                        }
                        catch (FileNotFoundException e)
                        {
                            _item.value = e.Message;
                        }
                    }

                    sheetData.AppendChild(getDataRow(item.Select(x => x.value), rowIndex));
                    
                    rowIndex++;
                }

                worksheet.Save();
            }
        }

        public class ExportItem
        {
            public object   value       { get; set; }
            public bool     isImage     { get; set; }
            public int?     imageWidthPx{ get; set; }
            public int      colNo       { get; set; }
            public int      rowNo       { get; set; }
        }

        #region Implement Print Image in Cell

        public class ExcelImagePathAttribute : Attribute
        {
            public ExcelImagePathAttribute()
            {
            }

            public ExcelImagePathAttribute(int aWidthPx)
            {
                WidthPx = aWidthPx;
            }

            public int? WidthPx { get; set; }
        }

        public static void AddImage(WorksheetPart worksheetPart, Stream imageStream, string imgDesc, int colNumber, int rowNumber, int? customWidth = null)
        {
            // We need the image stream more than once, thus we create a memory copy
            MemoryStream imageMemStream = new MemoryStream();
            imageStream.Position = 0;
            imageStream.CopyTo(imageMemStream);
            imageStream.Position = 0;

            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;

            Bitmap bm = new Bitmap(imageMemStream);
            var imagePart = drawingsPart.AddImagePart(GetImagePartTypeByBitmap(bm));
            imagePart.FeedData(imageStream);

            var width = bm.Width;
            var height = bm.Height;
            if (customWidth != null)
            {
                width = (int)customWidth;
                height = (int)(bm.Height * ((decimal) customWidth / bm.Width));
            }

            var extentsCx = width * (long)(914400 / bm.HorizontalResolution);
            var extentsCy = height * (long)(914400 / bm.VerticalResolution);
            bm.Dispose();

            var colOffset = 0;
            var rowOffset = 0;

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0
                ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1
                : 1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                    RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imgDesc },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }

        public static ImagePartType GetImagePartTypeByBitmap(Bitmap image)
        {
            if (ImageFormat.Bmp.Equals(image.RawFormat))
                return ImagePartType.Bmp;
            else if (ImageFormat.Gif.Equals(image.RawFormat))
                return ImagePartType.Gif;
            else if (ImageFormat.Png.Equals(image.RawFormat))
                return ImagePartType.Png;
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
                return ImagePartType.Tiff;
            else if (ImageFormat.Icon.Equals(image.RawFormat))
                return ImagePartType.Icon;
            else if (ImageFormat.Jpeg.Equals(image.RawFormat))
                return ImagePartType.Jpeg;
            else if (ImageFormat.Emf.Equals(image.RawFormat))
                return ImagePartType.Emf;
            else if (ImageFormat.Wmf.Equals(image.RawFormat))
                return ImagePartType.Wmf;
            else
                throw new Exception("Image type could not be determined.");
        }

        #endregion

        /// <summary> Get a typed list </summary>
        public static List<EntityType> getDataFromExcel<EntityType>(string aFileName, string aTableName = null) where EntityType : class, new() 
        {
            using (FileStream stream = File.Open(aFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var document = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart;
                workbookPart = document.WorkbookPart;
                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                var sheet = sheets.First(s => (aTableName == null || aTableName.Equals(s.Name))
                    && (s.State == null || s.State == SheetStateValues.Visible));

                var workSheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                var result = new List<EntityType>();

                var headerList = new Dictionary<string, string>();

                foreach (Row row in rows)
                {
                    var entity = Activator.CreateInstance<EntityType>();

                    var cellIdx = 0;
                    foreach (Cell cell in row)
                    {
                        cellIdx += 1;
                        var colName = GetColumnName(cell.CellReference) ?? cellIdx.ToString();
                        var val = GetCellValueFormatted(document.WorkbookPart, cell);
                        
                        if (row.RowIndex == 1)/*Header*/
                        {
                            var prop = entity.GetType().GetProperties()
                                .FirstOrDefault(x => x.GetCustomAttributes(false).Any(a => (a is DescriptionAttribute)
                                                                                        && (a as DescriptionAttribute).Description?.ToLower()?.Trim() == val?.ToLower()?.Trim())
                                                  || x.Name == val);
                            if (prop == null)
                            {
                                continue;
                            }
                            headerList[colName] = prop.Name;
                        }
                        else
                        {
                            string headerName;
                            if (!headerList.TryGetValue(colName, out headerName))
                            {
                                continue;
                            }
                            PropertyInfo propertyInfo = entity.GetType().GetProperty(headerName);
                            if (propertyInfo == null)
                            {
                                continue;
                            }
                            
                            var underlyingType = Nullable.GetUnderlyingType(propertyInfo.PropertyType);
                            if (underlyingType != null && string.IsNullOrWhiteSpace(val))
                            {
                                propertyInfo.SetValue(entity, null, null);
                                continue;
                            }

                            var t = underlyingType ?? propertyInfo.PropertyType;
                            if (t == typeof(decimal))
                            {
                                val = val.Replace(".", ",");
                            }

                            try
                            {
                                object convertedVal = Convert.ChangeType(val, t);
                                propertyInfo.SetValue(entity, convertedVal, null);
                            }
                            catch (Exception e)
                            {
                                throw new ArgumentException(string.Format("Cell {0}{1}, value \"{2}\" => Conver error {3}", colName, cellIdx, val, t.Name));
                            }
                        }
                    }

                    var props = entity.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    var r = props.All(p => p.GetValue(entity, null) == null);

                    if (r)
                    {
                        continue;
                    }

                    if (row.RowIndex > 1)
                    {
                        result.Add(entity);
                    }
                }

                return result;
            }
        }

        /// <summary> Get vell value </summary>
        private static string GetCellValueFormatted(WorkbookPart workbookPart, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
            {
                return null;
            }

            string value = "";
            if (cell.DataType == null) // number & dates
            {
                if (cell.StyleIndex == null)
                {
                    return cell.CellValue.InnerText;
                }

                int styleIndex = (int)cell.StyleIndex.Value;
                CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);
                uint formatId = cellFormat.NumberFormatId.Value;

                if (formatId == CustomStylesheet.Formats.DateShort || formatId == CustomStylesheet.Formats.DateLong)
                {
                    double oaDate;
                    if (double.TryParse(cell.InnerText, out oaDate))
                    {
                        value = DateTime.FromOADate(oaDate).ToShortDateString();
                    }
                }
                else
                {
                    value = cell.InnerText;
                }
            }
            else // Shared string or boolean
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.InnerText));
                        value = ssi.Text.Text;
                        break;
                    case CellValues.Boolean:
                        value = cell.CellValue.InnerText == "0" ? "false" : "true";
                        break;
                    default:
                        value = cell.CellValue.InnerText;
                        break;
                }
            }

            return value;
        }

        /// <summary> Get cell value (SpreadsheetDocument) </summary>
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return null;
            }
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText?.Trim();
            }
            else
            {
                return value?.Replace(".", ",");
            }
        }

        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            if (cellName == null)
            {
                return null;
            }
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
    }
}