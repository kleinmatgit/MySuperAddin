using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MySuperAddin._Common.Excel
{
    internal class ExcelTypeConversions
    {
        public static object ConvertObjectForOutput(object value)
        {
            //do not convert numerical types
            if ((value.GetType() == typeof(double)) ||
                (value.GetType() == typeof(Double)) ||
                (value.GetType() == typeof(int)) ||
                (value.GetType() == typeof(bool)))
                return value;
            //convert all other types to string: ExcelDna might not know how to display them (e.g enums)
            return value.ToString();
        }

        //converts a list into a range, padding it with #N/A's if needed
        public static object[,] ListToRange<T>(IList<T> input, bool vertical = true)
        {
            if (input.Count == 0) //nothing to do
                return ExcelAPI.ErrorNA2D;

            object[,] output;
            if(ExcelAPI.IsCalledFromSingleCell()) //do not resize
            {
                output = new object[(vertical ? input.Count : 1), (vertical ? 1 : input.Count)];
                for (int i = 0; i < input.Count; i++)
                    output[(vertical ? i : 0), (vertical ? 0 : i)] = input[i];
                return output;
            }

            bool adjust = (input.Count == 1 && !ExcelAPI.IsCalledFromSingleCell());
            int size = (adjust ? 2 : input.Count);
            output = new object[(vertical ? size : 2), (vertical ? 2 : size)];

            if(vertical) //display vector vertically
            {
                for(int row = 0; row < input.Count; row++)
                {
                    output[row, 0] = input[row];
                    output[row, 1] = ExcelAPI.ErrorNA;
                }
                if(adjust) //insert an additional column
                {
                    output[1, 0] = ExcelAPI.ErrorNA;
                    output[1, 1] = ExcelAPI.ErrorNA;
                }
            }
            else //display horizontally
            {
                for(int col = 0; col < input.Count; col++)
                {
                    output[0, col] = input[col];
                    output[1, col] = ExcelAPI.ErrorNA;
                }
                if(adjust) //insert additional row
                {
                    output[0, 1] = ExcelAPI.ErrorNA;
                    output[1, 1] = ExcelAPI.ErrorNA;
                }
            }
            return output;
        }

        public static List<object> RangeToListOfObjects(object[] range, bool preserveErrors = true)
        {
            //return a list populated with the values of the input array, skipping empty and missing cells
            return (from item in range
                    where !ExcelAPI.IsEmptyCell(item) && !ExcelAPI.IsMissingCell(item) && (!ExcelAPI.IsErrorCell(item) || preserveErrors)
                    select item).ToList();
        }

        public static List<string> RangeToListOfStrings(object[] range)
        {
            //handle the case where a null value is passed
            if (range == null)
                return new List<string>();
            //return a list populated with the values of the input array, skipping empty and missing cells
            return (from item in range
                    where (
                    !ExcelAPI.IsEmptyCell(item) &&
                    !ExcelAPI.IsMissingCell(item) &&
                    ExcelAPI.ThrowOnErrorCell(item))
                    select item.ToString()).ToList();
        }
    }
}
