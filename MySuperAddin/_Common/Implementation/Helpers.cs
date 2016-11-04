using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MySuperAddin._Common.Implementation
{
    //Various utility functions

    internal class Helpers
    {
        public static List<T> RemoveDuplicatesFromList<T>(List<T> input)
        {
            return RemoveDuplicatesFromCollection(input).ToList();
        }

        public static T[] ListToArray1D<T>(List<T> input)
        {
            var output = new T[input.Count];
            int row = 0;
            foreach(T item in input)
            {
                output[row] = item;
                row++;
            }
            return output;
        }

        public static List<T> SelectXmlNodesAsList<T>(string xml, string xpath)
        {
            XmlDocument doc = new XmlDocument();
            List<T> output = new List<T>();
            doc.LoadXml(xml);
            try
            {
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                RegisterXMLNamespaces(nsmgr);
                var iterator = doc.SelectNodes(xpath, nsmgr).GetEnumerator();
                if(iterator != null)
                {
                    while(iterator.MoveNext())
                    {
                        output.Add((T)Convert.ChangeType(((XmlNode)iterator.Current).InnerText, typeof(T)));
                    }
                }
            }
            catch(Exception e)
            {
                Logging.Logger.Log("Error applying XPath: " + e.Message);
            }
            return output;
        }


        #region Private methods

        private static void RegisterXMLNamespaces(XmlNamespaceManager nsmgr)
        {
            for (int count = 1; count < 100; count++) //assumes a maximum of 99 namespaces
            {
                string prefix = Config.ConfigController.GetConfigEntry("Namespace" + count.ToString() + ".Prefix");
                string uri = Config.ConfigController.GetConfigEntry("Namespace" + count.ToString() + ".URI");
                if ((prefix.Length > 0) && (uri.Length > 0))
                    nsmgr.AddNamespace(prefix, uri);
            }
        }

        private static IEnumerable<T> RemoveDuplicatesFromCollection<T>(IEnumerable<T> input)
        {
            var passedValues = new List<T>();
            foreach (T item in input)
                if (passedValues.Contains(item))
                    continue;
                else
                {
                    passedValues.Add(item);
                    yield return item;
                }
        }
        #endregion
    }
}
