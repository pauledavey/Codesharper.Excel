namespace ExcelAddIn.Functions.Functions.Test_Function
{
    using System;
    using ExcelDna.Integration;

    public class TestFunction
    {
        [ExcelFunction(Description = "Codesharper Test Function. Pass any string in to the function. Use for testing purposes only")]
        public static string CodesharperTestFunction(string userInputString)
        {
            return "You Entered " + userInputString;
        }
    }
}
