namespace ExcelAddIn.Functions.Functions.ColourCount
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Threading.Tasks;

    using ExcelDna.Integration;

    public class CellColourInformation
    {
        public Color colour { get; set; }
        public int colourInt { get; set; }
        public string colourName { get; set; }
        public int Counter { get; set; }
    }

    public class ColourCountFunction : IExcelAddIn
    {
        /// <summary>
        /// Dictionary for holding cell colours so we can correctly count and match them
        /// </summary>
        private static Dictionary<int, CellColourInformation> cellColours = new Dictionary<int, CellColourInformation>();

        /// <summary>
        /// This is the FUNCTON Extension main method
        /// </summary>
        /// <param name="myInput">The selection range from the spreadsheet</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Colour Count Function", IsMacroType = true)]
        public static int ColourCount([ExcelArgument(AllowReference = true, Name = "Range of cells to search in")] object myInputRange, [ExcelArgument(AllowReference = true, Name = "Cell that contains the colour we want to count for")] object myInputColour)
        {
            cellColours.Clear();
            ExcelReference myInputRef = myInputRange as ExcelReference;
            ParseCellCollection(myInputRef);

            dynamic appHandle = ExcelDnaUtil.Application;
            var refColourRange = appHandle.Range(XlCall.Excel(XlCall.xlfReftext, myInputColour, true));
            var selectedCellColour = Convert.ToInt32(refColourRange.Interior.Color);


            int result = 0;

            // Write back to the excel sheet!
            Parallel.ForEach(cellColours, (entry, state) =>
                    {
                        // write to the target cell!
                        CellColourInformation cellcolourInformation = entry.Value;

                        if (cellcolourInformation.colourInt == selectedCellColour)
                        {
                            result = cellcolourInformation.Counter;
                            state.Break();
                        }
                    });

            // If we got here then there were no colours found in the range
            return result;
        }

        /// <summary>
        /// Parse cells for colours
        /// </summary>
        /// <param name="selectionIn">Selected range In</param>
        [ExcelCommand(Description = "Parse Cells for Colours")]
        private static void ParseCellCollection ([ExcelArgument (AllowReference = true)] object selectionIn)
        {
            dynamic appHandle = ExcelDnaUtil.Application;
            var refRange = appHandle.Range(XlCall.Excel(XlCall.xlfReftext, selectionIn, true));

            // Test getting the colour of the cell. 
            foreach (var cell in refRange.Cells)
            {
                var cellColour = Convert.ToInt32(cell.Interior.Color);
                UpdateDictionary(cellColour);
            }
        }


        /// <summary>
        /// Recalculate Multiple Formulas in selected Range
        /// </summary>
        /// <param name="selectionIn"></param>
        private static void RecalculateFormulas([ExcelArgument(AllowReference = true)] dynamic selectionIn)
        {
            foreach (dynamic cellEntry in selectionIn)
            {
                // We have a cell to work with here
                if (cellEntry.HasFormula)
                {
                    selectionIn.Calculate();
                } 
            }
        }


        /// <summary>
        /// Update Dictionary
        /// </summary>
        /// <param name="cellColourStringInt">integer OLE colour code</param>
        private static void UpdateDictionary(int cellColourStringInt)
        {
            // if the cell is the default background colour of white then ignore it!
            if (ColourConverterFromOle(cellColourStringInt) == Color.White)
            {
                // ignore this
                return;
            }

            if (cellColours.ContainsKey(cellColourStringInt))
            {
                var updateInformation = cellColours[cellColourStringInt];
                updateInformation.Counter += 1;
                cellColours[cellColourStringInt] = updateInformation;
                return;
            }

            var newUpdateInformation = new CellColourInformation();
            newUpdateInformation.Counter = 1;
            newUpdateInformation.colour = ColourConverterFromOle(cellColourStringInt);
            newUpdateInformation.colourInt = cellColourStringInt;
            newUpdateInformation.colourName = newUpdateInformation.colour.Name;
            cellColours.Add(cellColourStringInt, newUpdateInformation);
            return;
        }

        /// <summary>
        /// Convert an OLE Colour to a System.Drawing.Color
        /// </summary>
        /// <param name="colourIn">Color as OLE Integer</param>
        /// <returns></returns>
        private static Color ColourConverterFromOle(int colourIn)
        {
           return ColorTranslator.FromOle(colourIn);
        }

        /// <summary>
        /// Convert a System.Drawing.Color to an OLE Colour Integer
        /// </summary>
        /// <param name="colourIn">Color as System.Drawing.Color</param>
        /// <returns></returns>
        private static int ColourConverterFromOle(Color colourIn)
        {
            return ColorTranslator.ToOle(colourIn);
        }

        public void AutoClose()
        {
            throw new NotImplementedException();
        }

        public void AutoOpen()
        {
            throw new NotImplementedException();
        }
    }
}
