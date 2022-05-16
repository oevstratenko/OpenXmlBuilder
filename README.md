# OpenXmlBuilder
Create, edit Word/Excel documents using OpenXml.<br />

## Process Word document
  It's functionality for replace some macro text to a text from your own data source provider (need to implement IMacroDataProvider).<br />
###### There are two types of macro: value or table
1. Value: [[macro:list0:value:ParentDocumentNo]] - single scalar value<br />
2. Table: [[macro:list1:table:Product{{width:40%}}|photoPath{{Image:100x100}}{{Caption:Photo}}|Price{{format:n2}}|Qnt{{Caption:Quantity}}|SaleDate{{format:dd.MM.yyyy}}|Summ{{format:n2}}]] - table, multiple columns<br />
  
###### Macro structure
1. [[]] - begin-end scope<br />
2. macro - macro type begin (maybe you extend implemetation, for example "image:", "barcode:" etc)<br />
3. list0 - it's key for identity your own data source from your own implemented data source provider<br />
4. value/table - type of macro (explaine above)<br />
5. ParentDocumentNo - name of field from data source (list0 is this case). It's may to have couple attributes:<br />
   - width - set in percent width of column inside table (only table type of macro)<br />
   - image - marks macro as path to image file that should to apears instead of macro. Has adjust fit of image in px (example 100x100)<br />
   - caption - table column name (only table type of macro)<br />
   - format - format of macro value (like c# toString(format), n2, dd.MM.yyyy etc)<br />
```
public void ProcessWord(string aNewFileName)
{
    var dataProvider = new DemoMacroDataProvider();
    dataProvider.init();
    
    new WordDocumentBuilder(dataProvider).Process($"{Directory.GetCurrentDirectory()}\\source\\template.docx", aNewFileName);
}
```
