using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySuperAddin._Common.Excel;

namespace MySuperAddin._Common.Depot
{
    internal class Depot
    {
        //Removes all objects from the depot
        public static void Clear()
        {
            _depot.Clear();
        }

        //Adds an object to the depot, overwriting any object with the same name
        public  static string Add(string name, object item)
        {
            int version = 1;
            if (name.Length == 0)
                name = ExcelAPI.GetCallingCellUniqueId();
            if (_depot.ContainsKey(name))
                version = _depot[name].Version + 1;
            _depot[name] = new VersionedItem { Version = version, Item = item };
            return name + ":" + version.ToString();
        }

        //Retrieve an object from the depot
        public static object GetAsObject(string name)
        {
            int pos = name.LastIndexOf(":");
            if (pos != -1)
                name = name.Substring(0, pos);
            return (_depot.ContainsKey(name) ? _depot[name].Item : null);
        }

        //Retrieve an object from the depot and cast it to type T
        public static T Get<T>(string name) where T :class
        {
            object itemAsObject = GetAsObject(name);
            if ((itemAsObject != null) && (itemAsObject.GetType() == typeof(T)))
                return (T)itemAsObject;
            return null;
        }
        
        //Retrieve a string representation of an object in the depot
        public static string GetAsString(string name)
        {
            object itemAsObject = GetAsObject(name);
            if (itemAsObject != null)
                return itemAsObject.ToString();
            return "";
        }

        //Number of entries in the depot
        public static int Count
        {
            get { return _depot.Count; }
        }

        #region Private fields and methods

        //object container
        private struct VersionedItem
        {
            public int Version { get; set; }
            public object Item { get; set; }
        }

        //private constructor
        private Depot() { }

        //the object store
        private static readonly Dictionary<string, VersionedItem> _depot = new Dictionary<string, VersionedItem>();

        #endregion
    }
}
