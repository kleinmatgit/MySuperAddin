using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace MySuperAddin.SEO.ExcelInterface
{
    public class SEOExcelInterface
    {
        [ExcelFunction(Description = "Test SEO Excel Interface")]
        public static string SEOTest()
        {
            return "SEO Excel Interface is working fine.";
        }
    }
}
