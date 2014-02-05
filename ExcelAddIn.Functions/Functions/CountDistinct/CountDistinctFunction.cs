using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;

namespace ExcelAddIn.Functions.Functions.CountDistinctCellValues
{
    public class CountDistinctFunction
    {
        private static Dictionary<string,int> distinctList =new Dictionary<string, int>();

        [ExcelFunction(Description = "Count Distinct Cell Values", IsMacroType = true)]
        public static string CountDistinct([ExcelArgument(AllowReference = true, Name = "Range of cells to search in")] object myInputRange, [ExcelArgument(AllowReference = true, Name = "Show Count")] bool showCount)
        {
            distinctList.Clear();
            dynamic appHandle = ExcelDnaUtil.Application; 
            ExcelReference myInputRef = myInputRange as ExcelReference;
            ParseCellCollection(myInputRef);

            if (distinctList.Count == 0)
            {
                return "No Distinct Values were found";
            }

            if (showCount)
            {
                return string.Join("   ", distinctList.Select(x => x.Key + "=" + x.Value).ToArray());
            }

                return string.Join("   ", distinctList.Select(x => x.Key).ToArray());

        }

        /// <summary>
        /// Parse cells for distinct values
        /// </summary>
        /// <param name="selectionIn">Selected range In</param>
        [ExcelCommand(Description = "Parse Cells for Distinct values")]
        private static void ParseCellCollection([ExcelArgument(AllowReference = true)] object selectionIn)
        {
            dynamic appHandle = ExcelDnaUtil.Application;
            var refRange = appHandle.Range(XlCall.Excel(XlCall.xlfReftext, selectionIn, true));

            // Test getting the colour of the cell. 
            foreach (var cell in refRange.Cells)
            {
                UpdateDictionary(cell.Value2.ToString());
            }
        }

        /// <summary>
        /// Update Dictionary
        /// </summary>
        /// <param name="cellValueIn">integer OLE colour code</param>
        private static void UpdateDictionary(string cellValueIn)
        {

            if (string.IsNullOrEmpty(cellValueIn))
            {
                // ignore blank cells
                return;
            }

            if (distinctList.ContainsKey(cellValueIn))
            {
                distinctList[cellValueIn] += 1;
                return;
            }

            distinctList.Add(cellValueIn, 1);
            return;
        }

    }
}
