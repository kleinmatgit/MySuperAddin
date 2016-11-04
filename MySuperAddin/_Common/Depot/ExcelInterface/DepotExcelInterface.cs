using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace MySuperAddin._Common.Depot.ExcelInterface
{
    public class DepotExcelInterface
    {
        private const string CATEGORY = "Depot";

        [ExcelFunction(Description ="Test Depot Excel Interface")]
        public static string TestDepot()
        {
            return Depot.Add("testOject", new object());
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description = "Displays a textual representation of the specifie object",
            Name ="DepotShowContent")]
        public static string ShowContents(
            [ExcelArgument(
                Description="Object handle",
                Name = "Handle")] string handle)
        {
            return Depot.GetAsString(handle);
        }

        [ExcelFunction(
            Category = CATEGORY,
            Description ="Deletes all objects stored in the depot",
            Name = "DepotClear")]
        public static int ClearDepot()
        {
            Depot.Clear();
            return Depot.Count;
        }
    }
}
