# OpenXmlBuilder
Create, edit Word/Excel documents using OpenXml.

## Process Word document
  It's functionality for replace some macro text to a text from your own data source provider (need to implement IMacroDataProvider).
###### There are two types of macro: value or table
1. Value: [[macro:list0:value:ParentDocumentNo]] - single scalar value
2. Table: [[macro:list1:table:Product{{width:40%}}|photoPath{{Image:100x100}}{{Caption:Photo}}|Price{{format:n2}}|Qnt{{Caption:Quantity}}|SaleDate{{format:dd.MM.yyyy}}|Summ{{format:n2}}]] - table, multiple columns
  
###### Macro structure
1. [[]] - begin-end scope
2. macro - macro type begin (maybe you extend implemetation, for example "image:", "barcode:" etc)
3. list0 - it's key for identity your own data source from your own implemented data source provider
4. value/table - type of macro (explaine above)
5. ParentDocumentNo - name of field from data source (list0 is this case). It's may to have couple attributes:
   - width - set in percent width of column inside table (only table type of macro)
   - image - marks macro as path to image file that should to apears instead of macro. Has adjust fit of image in px (example 100x100)
   - caption - table column name (only table type of macro)
   - format - format of macro value (like c# toString(format), n2, dd.MM.yyyy etc)
```
public void ProcessWord(string aNewFileName)
{
    var dataProvider = new DemoMacroDataProvider();
    dataProvider.init();
    
    new WordDocumentBuilder(dataProvider).Process($"{Directory.GetCurrentDirectory()}\\source\\template.docx", aNewFileName);
}
```

## Process Excel document
  Create and edit Excel document. There is a ExcelSheetBuilder that accumulate commands and parameters to build sheet of excel document.
###### Exists mathods:
1. AppendText - append a simple text to current cell. Has two params: text that will be append, CellStyle - style of cell and content.
   - CellStyle - style of cell and content - has poperties:
     - BorderStyle - boder line style, enum, native OpneXml
     - Height - height of cell
     - MergeCnt - cells count for merge in one
     - BackColorRgb - back color of cell
     - Alignment - text alignment
     - Font - font style, native OpneXml
2. AppendLine - make new line break. Has no params
3. AppendTable - append a table from IEnumerable. Has two params: data that will be append, TableSetting - style of table.
   - TableSetting - style of table - has poperties:
     - HeaderCellStyle - header cells style, type CellStyle
     - DataCellStyle - table body cells style, type CellStyle
     - HideHeader - is need to hide header of table, true/false (false by default)
     - ExportedColumns - custom columns names for export from datasource (null by default - export all)

__Can export image as column data type where value as path of file. You should to mark that field by ExcelHelper.ExcelImagePath attribute__

```
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
 ```
