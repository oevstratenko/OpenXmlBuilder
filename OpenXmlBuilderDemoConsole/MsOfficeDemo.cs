using OpenXmlBuilder;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;

namespace OpenXmlBuilderDemoConsole
{
    internal class MsOfficeDemo
    {
        /// <summary> Create docx demo file </summary>
        public void ProcessWord(string aNewFileName)
        {
            var dataProvider = new DemoMacroDataProvider();
            dataProvider.init();

            new WordDocumentBuilder(dataProvider).Process($"{Directory.GetCurrentDirectory()}\\source\\template.docx", aNewFileName);
        }

        /// <summary> Create xlsx demo file </summary>
        public void CreateExcel(string aFileName)
        {
            var list = Enumerable.Range(0, 100).Select(x => new ListItem { Id = x, Name = $"Name {x}", Description = $"Description {x}" });            
            list.exportToExcel(aFileName, "Sheet1");

            var esb = new ExcelSheetBuilder();

            esb.AppendText("MsOfficeDemo: using OpenXmlBuilder demonstration", new ExcelSheetBuilder.CellPolygon.CellStyle
            {
                MergeCnt    = 14,
                Alignment   = new Alignment { Horizontal = HorizontalAlignmentValues.Center },
                Font        = new Font { Bold = new Bold() }
            }).AppendLine()
            .AppendLine()
            .AppendText("Some demonstration table");

            esb.AppendTable(list, new ExcelSheetBuilder.TableSetting
            {
                HeaderCellStyle = new ExcelSheetBuilder.CellPolygon.CellStyle
                {
                    BorderStyle     = BorderStyleValues.Thin,
                    Height          = 40,
                    BackColorRgb    = System.Drawing.Color.FromArgb(252, 213, 180).ToHex(),
                    Alignment       = new Alignment { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }
                }
            });
            
            esb.SaveToFile(aFileName, "Sheet2", false);
        }
    }

    public class ListItem
    {
        public int      Id          { get; set; }
        public string   Name        { get; set; }
        public string   Description { get; set; }
        [ExcelHelper.ExcelImagePath(180)]
        public string Image
        {
            get
            {
                if (Id % 10 != 0)
                {
                    return string.Empty;
                }
                return $"{Directory.GetCurrentDirectory()}\\Source\\ms-.net-framework.jpg";
            }
            set { }
        }
    }

    internal static class Ext
    {
        /// <summary> System.Drawing.Color convert to Hex string </summary>
        internal static string ToHex(this System.Drawing.Color color)
        {
            return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }
    }
}
