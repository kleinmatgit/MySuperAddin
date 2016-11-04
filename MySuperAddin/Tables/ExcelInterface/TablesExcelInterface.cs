using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using MySuperAddin._Common.Excel;

namespace MySuperAddin.Tables.ExcelInterface
{
    public class TablesExcelInterface
    {
        private const string CATEGORY = "Tables";

        [ExcelFunction(
            Name ="TablesConcatenateRange",
            Description ="Concatenate the values of the cells in range.",
            Category = CATEGORY)]
        public static object ConcatenateRange(
            [ExcelArgument(
                Description = "input range that we want to concatenate", 
                Name = "Input Range")] object[] input, 
            [ExcelArgument(
                Description = "separator used between each value of the range in the output string", 
                Name = "Separator")] string separator)
        {
            var inputList = ExcelTypeConversions.RangeToListOfStrings(input);
            return string.Join(separator, inputList);
        }
    }
}
