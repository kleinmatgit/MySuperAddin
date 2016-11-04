using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Text.RegularExpressions;

namespace MySuperAddin._Common.Excel.ExcelInterface
{
    public class ExcelAPIExcelInterface
    {
        private const string CATEGORY = "Excel";

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Count the number of cells containing non-error values in a vector",
            Name = "ExcelCountNonErrorCells")]
        public static int CountNonErrorCells(
            [ExcelArgument(
                Description = "Input range",
                Name = "Range")] object[] input)
        {
            var inputList = ExcelTypeConversions.RangeToListOfObjects(input, false);
            return inputList.Count;
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the formula in the first cell of the specified range",
            IsMacroType = true,
            Name = "ExcelGetFormula")]
        public static object GetFormula(
            [ExcelArgument(
                AllowReference = true,
                Description = "Input range",
                Name = "Range")] object range)
        {
            return XlCall.Excel(XlCall.xlfGetFormula, range);
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the type of the value passed to the function",
            IsMacroType = true,
            Name = "ExcelGetArgumentType")]
        public static string GetArgumentType(
            [ExcelArgument(
                AllowReference = true,
                Description = "Argument for which we want to check the type",
                Name = "Argument")] object arg,
            [ExcelArgument(
                Description = "Excel references will be ignored if set to true (default)",
                Name = "Ignore References Flag")] bool ignoreReferences = true)
        {
            return ExcelAPI.GetArgumentType(arg, ignoreReferences);
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the name of the calling workbook",
            Name = "ExcelGetWorkbookName")]
        public static string GetWorkbookName()
        {
            return ExcelAPI.GetCallingWorkbookName();
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the name of the calling worksheet",
            Name = "ExcelGetWorksheetName")]
        public static string GetWorksheetName()
        {
            return ExcelAPI.GetCallingWorksheetName();
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the index of the calling workbook",
            Name = "ExcelGetWorkbookIndex")]
        public static int GetWorkbookIndex()
        {
            return ExcelAPI.GetCallingWorkbookIndex();
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the index of the calling worksheet",
            Name = "ExcelGetWorksheetIndex")]
        public static int GetWorksheetIndex()
        {
            return ExcelAPI.GetCallingWorksheetIndex();
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the full path of the calling workbook",
            Name = "ExcelGetWorkbookPath")]
        public static string GetWorkbookPath()
        {
            var folder = ExcelAPI.GetCallingWorkbookFolder();
            return System.IO.Path.Combine(folder, ExcelAPI.GetCallingWorkbookName());
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the version of the current Excel session",
            Name = "ExcelGetVersion")]
        public static object GetExcelVersion()
        {
            return ExcelDnaUtil.ExcelVersion;
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the process id of the current Excel session",
            Name = "ExcelGetProcessId")]
        public static int GetProcessId()
        {
            System.Diagnostics.Process currentProcess = System.Diagnostics.Process.GetCurrentProcess();
            return currentProcess.Id;
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Returns the style of the cells as text",
            IsMacroType = true,
            Name = "ExcelGetCellStyle")]
        public static object GetCellStyle(
            [ExcelArgument(
                AllowReference = true,
                Description = "Cell for which we want to check the style",
                Name = "Argument")] object arg)
        {
            return ExcelAPI.GetCellStyleName(arg);
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Searches input string for all occurences of the specified regular expression",
            Name = "ExcelFindRegExMatches")]
        public static object[,] FindRegExMatches(
            [ExcelArgument(
                AllowReference = true,
                Description = "String for which we want to search regular expression",
                Name = "Input String")] string input,
            [ExcelArgument(
                AllowReference = true,
                Description = "Regular expression we are trying to match",
                Name = "Regular Expression")] string regex)
        {
            MatchCollection matches = Regex.Matches(input, regex);
            List<string> list = (from Match m in matches select m.Groups[1].Value).ToList();
            return ExcelTypeConversions.ListToRange<string>(list, true);
        }

    }
}
