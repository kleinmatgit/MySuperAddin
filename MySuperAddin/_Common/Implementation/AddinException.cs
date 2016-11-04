using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MySuperAddin._Common.Implementation
{
    internal class AddInException : Exception
    {
        public AddInException() : base() { }
        public AddInException(string message) : base(message) { }
        public AddInException(string message, Exception inner) : base(message, inner) { }
    }
}
