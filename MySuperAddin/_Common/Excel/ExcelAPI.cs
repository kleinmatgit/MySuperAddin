using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using MySuperAddin._Common.Implementation;


namespace MySuperAddin._Common.Excel
{
    //
    //Various useful Excel API wrappers
    //For every function in this class we indicate whether macros permissions are required,
    //i.e. whether the functon should be registered with IsMacroType set to true.
    //

    internal class ExcelAPI
    {
        public enum GetCellTypes
        {
            ReturnStyle = 40,
            ReturnWorkbookName = 66
        };

        public static object[,] ErrorNA2D = new object[,] { { ExcelError.ExcelErrorNA } };
        public static object[] ErrorNA1D = new object[] { ExcelError.ExcelErrorNA };
        public static object ErrorNA = ExcelError.ExcelErrorNA;

        public static Application TheApplication = (Application)ExcelDnaUtil.Application;

        //does not require macro permissions
        public static string GetCallingWorkbookAndWorksheetNames()
        {
            return (string)XlCall.Excel(XlCall.xlSheetNm, GetCaller());
        }

        //does not require macro permissions
        public static string GetCallingCellUniqueId()
        {
            return GetCallingWorkbookAndWorksheetNames() + ";" + GetCallingCellAddress();
        }

        //does not require macro permissions
        public static string GetCallingWorkbookName()
        {
            var fullName = GetCallingWorkbookAndWorksheetNames();
            int closingBracket = fullName.IndexOf("]");
            if (closingBracket == -1) return fullName;
            return fullName.Substring(1, closingBracket - 1);
        }

        //does not require macro permissions
        public static int GetCallingWorkbookIndex()
        {
            int index = 0;
            foreach (Workbook workbook in TheApplication.Workbooks)
            {
                index++;
                if (workbook.Name == GetCallingWorkbookName())
                    return index;
            }
            return -1;
        }

        //does not require macro permissions
        public static string GetCallingWorksheetName()
        {
            var fullName = GetCallingWorkbookAndWorksheetNames();
            int closingBracket = fullName.IndexOf("]");
            if (closingBracket == -1) return fullName;
            return fullName.Substring(closingBracket + 1);
        }

        //does not require macro permissions
        public static int GetCallingWorksheetIndex()
        {
            int index = 0;
            foreach (Worksheet worksheet in GetWorksheetsInCallingWorkbook())
            {
                index++;
                if (worksheet.Name == GetCallingWorksheetName())
                    return index;
            }
            return -1;
        }

        //requires macro permissions
        public static string GetCallingWorkbookFolder()
        {
            var sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, GetCaller());
            return (string)XlCall.Excel(XlCall.xlfGetDocument, 2, sheetName);
        }

        //does not require macro permissions
        public static string GetCallingCellAddress()
        {
            return (GetCaller().ColumnFirst + 1) + ";" + (GetCaller().RowFirst + 1);
        }

        //does not require macro permissions
        public static Sheets GetWorksheetsInCallingWorkbook()
        {
            Workbook wbk = TheApplication.Workbooks[GetCallingWorkbookName()];
            return wbk.Worksheets;
        }

        //does not require macro permissions
        public static bool IsErrorCell(object input)
        {
            return (input.GetType() == typeof(ExcelError));
        }

        //does not require macro permissions
        public static bool IsEmptyCell(object input)
        {
            return (input.GetType() == typeof(ExcelEmpty));
        }

        //does not require macro permissions
        public static bool IsMissingCell(object input)
        {
            return (input.GetType() == typeof(ExcelMissing));
        }

        //does not require macro permissions
        public static bool ThrowOnErrorCell(object input)
        {
            if (input.GetType() == typeof(ExcelError))
                throw (new AddInException("Error in input"));
            return true;
        }

        //does not require macro permissions
        public static ExcelReference GetCaller()
        {
            return (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        }

        //does not require macro permissions
        public static bool IsCalledFromSingleCell()
        {
            var rows = GetCaller().RowLast - GetCaller().RowFirst + 1;
            var columns = GetCaller().ColumnLast - GetCaller().ColumnFirst + 1;
            return (rows == 1 && columns == 1);
        }

        //requires macro permissions
        public static string GetCellStyleName(object arg)
        {
            return (string)XlCall.Excel(XlCall.xlfGetCell, (int)GetCellTypes.ReturnStyle, arg);
        }

        //does not require macro permissions
        public static string GetArgumentType(object arg, bool ignoreReferences)
        {
            if (arg is double)
                return "Double";
            else if (arg is string)
                return "String";
            else if (arg is bool)
                return "Boolean";
            else if (arg is ExcelError)
                return "ExcelError: " + arg.ToString();
            else if (arg is object[,])
                return string.Format("Array[{0},{1}]", ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
            else if (arg is ExcelMissing)
                return "Missing";
            else if (arg is ExcelEmpty)
                return "Empty";
            else if (arg is ExcelReference)
            {
                var range = (ExcelReference)arg;
                if (ignoreReferences)
                    return GetArgumentType(range.GetValue(), true);
                else
                    return "Reference: " + XlCall.Excel(XlCall.xlfReftext, range, true)
                        + ", " + GetArgumentType(range.GetValue(), false);
            }
            else
                return "Unknown Type";
        }
    }
}
