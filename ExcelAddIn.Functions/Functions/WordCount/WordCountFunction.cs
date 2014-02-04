using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelAddIn.Functions.Functions
{
    public static class WordCountFunction
    {

        /// <summary>
        /// Dictionary that is used to store the WORD and word counts.
        /// </summary>
        private static Dictionary<string, int> wordCounts = new Dictionary<string, int>();


        /// <summary>
        /// Main Function To Count Instances of a WORD
        /// </summary>
        /// <param name="myInputRange">In Range from Worksheet</param>
        /// <param name="myInputWord">Word to check for</param>
        /// <param name="exactMatch">Exact Match (Inc. Case) or contains (ignore case)</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Word Count Function", IsMacroType = true)]
        public static int WordCount([ExcelArgument(AllowReference = true)] object myInputRange, [ExcelArgument(AllowReference = true)] object myInputWord, [ExcelArgument(AllowReference = true)] bool exactMatch)
        {
            wordCounts.Clear();
            ExcelReference myInputRef = myInputRange as ExcelReference;
            ParseCellCollection(myInputRef);

            dynamic appHandle = ExcelDnaUtil.Application;
            var refWordRange = appHandle.Range(XlCall.Excel(XlCall.xlfReftext, myInputWord, true));
            var selectedWord = refWordRange.Value2.ToString();

            int result = 0;

            if (exactMatch)
            {
                Parallel.ForEach(wordCounts, (entry, state) =>
                {
                    // write to the target cell!
                    var ret = string.Compare(entry.Key, selectedWord, StringComparison.CurrentCulture);
                    if (ret == 0)
                    {
                        result += entry.Value;
                    }
                });
            }
            else
            {
                Parallel.ForEach(wordCounts, (entry, state) =>
                {
                    if (entry.Key.IndexOf(selectedWord, StringComparison.CurrentCultureIgnoreCase) >= 0)
                    {
                        result += entry.Value;
                    }
                });
            }

            // If we got here then there were no words found in the range
            return result;
        }


        /// <summary>
        /// Parse cells for words
        /// </summary>
        /// <param name="selectionIn">Selected range In</param>
        [ExcelCommand(Description = "Parse Cells for Words")]
        private static void ParseCellCollection([ExcelArgument(AllowReference = true)] object selectionIn)
        {
            dynamic appHandle = ExcelDnaUtil.Application;
            var refRange = appHandle.Range(XlCall.Excel(XlCall.xlfReftext, selectionIn, true));

            // Test getting the word of the cell. 
            foreach (var cell in refRange.Cells)
            {
                var cellWord = cell.Value2;
                UpdateDictionary(cellWord);
            }
        }


        /// <summary>
        /// Update Dictionary
        /// </summary>
        /// <param name="cellwordStringInt">integer OLE word code</param>
        private static void UpdateDictionary(string cellWord)
        {
            // if the cell is the default background word of white then ignore it!
            if (string.IsNullOrEmpty(cellWord))
            {
                // ignore this
                return;
            }

            if (wordCounts.ContainsKey(cellWord))
            {
                wordCounts[cellWord] += 1;
                return;
            }

            wordCounts.Add(cellWord, 1);
            return;
        }
    }
}
