using System;
using System.IO;

namespace OpenXmlBuilder
{
    /// <summary> Process word document: replace macro with real values from IMacroDataProvider </summary>
    public class WordDocumentBuilder
    {
        public readonly IMacroDataProvider MacroDataProvider;
        public WordDocumentBuilder(IMacroDataProvider aDataProvider)
        {
            MacroDataProvider = aDataProvider;
        }

        private const string macroStart = "[[";
        private const string macroEnd = "]]";

        private const string macro1 = macroStart + "macro:";

        public void Process(string aSourcePath, string aResultPath)
        {
            File.Copy(aSourcePath, aResultPath, true);

            var pList = WordHelper.GetParagraphs(aResultPath);

            var macroBC = new Macro(MacroDataProvider);

            var pNo = 1;
            foreach (var item in pList)
            {
                int index = 0;

                while (index < item.Length)
                {
                    //[[macro:
                    if (item.Substring(index).StartsWith(macro1, StringComparison.InvariantCultureIgnoreCase))
                    {
                        var macro = item.Substring(index + macro1.Length, item.IndexOf(macroEnd, index, StringComparison.InvariantCultureIgnoreCase) - index - macro1.Length);

                        object value = macroBC.GetValue(macro);

                        WordHelper.ReplaceText(aResultPath, macroStart, macroEnd, value, pNo);
                    }
                    index++;
                }

                pNo++;
            }
        }
    }
}