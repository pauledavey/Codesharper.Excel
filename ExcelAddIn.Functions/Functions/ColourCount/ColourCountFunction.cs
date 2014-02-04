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
        private static Dictionary<int, CellColourInformation> cellColours = new Dictionary<int, CellColourInformation>();

        ///// <summary>
        ///// Recalculate Method handler for recalcing selected formula cells
        ///// </summary>
        ///// <param name="myInputRange">Selection in the sheet</param>
        //[ExcelCommand(Description = "Recalculate if it is one of our formulas", MenuName = "Colour Count", MenuText = "Recalc Selection")]
        //public static void Recalculate([ExcelArgument(AllowReference = true)] object myInputRange)
        //{
        //    dynamic appHandle = ExcelDnaUtil.Application;
        //    dynamic worksheetHandle = appHandle.ActiveSheet;


        //    if (appHandle.Selection == null)
        //    {
        //        return;
        //    }

        //    RecalculateFormulas(appHandle.Selection);
        //}


        /// <summary>
        /// This is the FUNCTON Extension main method
        /// </summary>
        /// <param name="myInput">The selection range from the spreadsheet</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Colour Count Function", IsMacroType = true)]
        public static int ColourCount([ExcelArgument(AllowReference = true)] object myInputRange, [ExcelArgument(AllowReference = true)] object myInputColour)
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
