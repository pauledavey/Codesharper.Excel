namespace ExcelAddIn.Functions.Functions.Test_Function
{
    using System;
    using ExcelDna.Integration;

    public class TestFunction
    {
        [ExcelFunction(Description = "Codesharper Test Function")]
        public static string CodesharperTestFunction(string userInputString)
        {
            return "You Entered " + userInputString;
        }
    }
}
