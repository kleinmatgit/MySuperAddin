using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySuperAddin._Common.Logging;
using Microsoft.Win32;
using MySuperAddin._Common.Excel;

namespace MySuperAddin._Common.Implementation
{
    internal class WinRegistry
    {
        public static object GetDLLPathFromProgId(string progId)
        {
            string classId = GetClassIdFromProgId(progId);
            if (classId != null)
            {
                string path = GetDLLPathFromClassId(classId);
                if (path != null)
                    return path;
            }
            return ExcelAPI.ErrorNA;
        }

        #region Private members

        private static string GetDLLPathFromClassId(string classId)
        {
            var regPath = @"\CLSID\" + classId + @"\InProcServer32\";
            return GetDefaultRegistryValue(Registry.ClassesRoot, regPath);
        }

        private static string GetClassIdFromProgId(string progId)
        {
            var regPath = progId + @"\CLSID\";
            return GetDefaultRegistryValue(Registry.ClassesRoot, regPath);
        }

        private static string GetDefaultRegistryValue(RegistryKey rootKey, string regPath)
        {
            try
            {
                var regKey = rootKey.OpenSubKey(regPath);
                if (regKey != null)
                {
                    return (string)regKey.GetValue("");
                }
                else
                {
                    Logger.Log("No default value returned for registry key " + regPath);
                }
            }
            catch (Exception e)
            {
                Logger.Log("Error accessing registry key " + regPath + ": " + e.Message);
            }
            return null;
        }

        #endregion
    }
}
