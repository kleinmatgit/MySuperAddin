using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace MySuperAddin._Common.Logging.ExcelInterface
{
    public class LoggingExcelInterface
    {
        private const string CATEGORY = "Logging";

        [ExcelFunction(
            Name = "LoggingShowTraceWindow", 
            Description ="Display or hide the trace window.",
            Category = CATEGORY)]
        public static bool ShowTraceWindow(
            [ExcelArgument(Description = "true/false to show/hide")] bool show)
        {
            Logger.ShowTraceWindow(show);
            return show;
        }
    }
}
