using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using DataTable = System.Data.DataTable;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;

namespace OpenXmlBuilder
{
    /// <summary> Helper for process docx (using OpenXml) </summary>
    public static class WordHelper
    {
        /// <summary> To search and replace content in a document part. </summary>
        /// <param name="document">File Path</param>
        public static string[] GetParagraphs(string document)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true))
            {
                return wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().Select(x => x.InnerText).ToArray();
            }
        }

        /// <summary> Replace text inside document paragraph </summary>
        /// <param name="document"> path to docx file </param>
        /// <param name="startText">substring for text replacement begin</param>
        /// <param name="endText">substring for text replacement end</param>
        /// <param name="newObject">new value</param>
        /// <param name="paragraphNo"> order number of document paragraph </param>
        public static void ReplaceText(string document, string startText, string endText, object newObject, int paragraphNo)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true))
            {
                Paragraph paragraph = wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().Skip(paragraphNo-1).Take(1).FirstOrDefault();
                if (paragraph == null)
                {
                    return;
                }

                var effects = new List<Run>();
                var isStartFinded = false;
                var isEndFinded = false;
                foreach (Run run in paragraph.Descendants<Run>().ToArray())
                {
                    if (run.InnerText.Contains(startText))
                    {
                        isStartFinded = true;
                    }

                    if (isStartFinded)
                    {
                        effects.Add(run);

                        if (run.InnerText.Contains(endText))
                        {
                            isEndFinded = true;
                            break;
                        }
                    }
                }

                if (!isEndFinded)
                {
                    return;
                }

                if (newObject is DataTable)
                {
                    var dt = newObject as DataTable;
                    var t = DataTableToOpenXmlTable(dt, dt.Columns.Count <= 1, wordprocessingDocument.MainDocumentPart);
                    var prop = effects.FirstOrDefault().RunProperties;

                    var rows = t.Elements<TableRow>().ToList();
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<TableCell>().ToList();
                        foreach (var run in cells.Select(cell => cell.Descendants<Run>().FirstOrDefault()))
                        {
                            var _prop = prop.Clone() as RunProperties;
                            if (rows.FirstOrDefault() == row && cells.Count() > 1)
                            {
                                _prop.Bold = new Bold();
                            }
                            run.RunProperties = _prop;
                        }
                    }

                    effects.FirstOrDefault().InsertBeforeSelf(new Run(t));
                }
                else
                {
                    var elFirst = effects.FirstOrDefault();
                    var beforeText = elFirst.InnerText.Substring(0, elFirst.InnerText.IndexOf(startText, StringComparison.InvariantCultureIgnoreCase));

                    var elLast = effects.LastOrDefault();
                    var afterText = elLast.InnerText.Substring(elLast.InnerText.IndexOf(endText, StringComparison.InvariantCultureIgnoreCase) + endText.Length, elLast.InnerText.Length - elLast.InnerText.IndexOf(endText, StringComparison.InvariantCultureIgnoreCase) - endText.Length);

                    var t = new Text(beforeText + newObject + afterText)
                    {
                        /*not trim freespace*/
                        Space = SpaceProcessingModeValues.Preserve
                    };

                    var newPar = new Run(t) { RunProperties = (elFirst.RunProperties?.Clone() ?? new RunProperties()) as RunProperties };
                    elFirst.InsertBeforeSelf(newPar);
                }

                foreach (Run run in effects)
                {
                    run.Remove();
                }
            }
        }
        
        private static Table DataTableToOpenXmlTable(DataTable aTable, bool aIsList = false, MainDocumentPart aMainPart = null)
        {
            var table = new Table();

            var tblProps = new TableProperties();
            var tableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
            var tableStyle = new TableStyle { Val = aIsList ? string.Empty : "TableGrid" };
            tblProps.Append(tableStyle, tableWidth);

            table.Append(tblProps);

            TableRow row;

            if (!aIsList)
            {
                row = new TableRow();
                foreach (DataColumn col in aTable.Columns)
                {
                    var cell = new TableCell();
                    cell.Append(new Paragraph(new Run(new Text(string.IsNullOrEmpty(col.Caption) ? col.ColumnName : col.Caption))));

                    var width = col.ExtendedProperties["width"]?.ToString();
                    cell.Append(new TableCellProperties(new TableCellWidth { Type = string.IsNullOrWhiteSpace(width) ? TableWidthUnitValues.Auto : TableWidthUnitValues.Pct, Width = width}));
                    row.Append(cell);
                }
                table.Append(row);
            }

            //columns names array with paths to image file, wich must be replace with this image
            var imageArray = aTable.ExtendedProperties["imageArray"] as ImageColumn[] ?? new ImageColumn[0];

            foreach (DataRow tRow in aTable.Rows)
            {
                row = new TableRow();
                foreach (DataColumn col in aTable.Columns)
                {
                    var cell = new TableCell();

                    var run = new Run();

                    var imageCol = imageArray.FirstOrDefault(x => x.Name == col.ColumnName);

                    if (imageCol == null)//text
                    {
                        cell.Append(new Paragraph(AddTextToRun(run, tRow[col.ColumnName].ToString())));
                    }
                    else//image
                    {
                        var p = new Paragraph();
                        cell.Append(p);

                        var fn = tRow[col.ColumnName].ToString();

                        ImagePart imagePart = aMainPart.AddImagePart(ImagePartType.Jpeg);
                        using (var stream = new FileStream(fn, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            imagePart.FeedData(stream);
                        }

                        Image img = Image.FromFile(fn);

                        AddImageToBody(p, aMainPart.GetIdOfPart(imagePart), imageCol.Width ?? img.Width, imageCol.Height ?? img.Height);
                    }

                    cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                    row.Append(cell);
                }
                table.Append(row);
            }

            return table;
        }

        /// <summary> append text </summary>
        private static Run AddTextToRun(Run run, string text)
        {
            var textArray = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var first = true;

            foreach (string line in textArray)
            {
                if (!first)
                {
                    run.AppendChild(new Break());
                }

                first = false;

                run.AppendChild(new Text { Text = line });
            }

            return run;
        }

        private static void AddImageToBody(Paragraph paragraph, string relationshipId, int aWidth, int aHeight)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = aWidth * 9525, Cy = aHeight * 9525 },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = aWidth * 9525, Cy = aHeight * 9525 }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         ) { Preset = A.ShapeTypeValues.Rectangle }))
                             ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            paragraph.AppendChild(new Run(element));
        }

        /// <summary> Image column </summary>
        public class ImageColumn
        {
            /// <summary> Column name </summary>
            public string Name { get; set; }
            /// <summary> Image width, px </summary>
            public int? Width { get; set; }
            /// <summary> Image height, px </summary>
            public int? Height { get; set; }
        }
    }
}