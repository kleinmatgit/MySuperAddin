using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MySuperAddin._Common.Logging
{
    public class Logger
    {
        public enum LogLevel
        {
            Debug,
            Info,
            Warning,
            Error,
            Critical
        };

        public static void Log(string msg, LogLevel level = LogLevel.Warning)
        {
#if DEBUG
            //in debug mode log everything
            WriteMessageToTraceWindow(msg);
#else
            //in release mode do not log anything below warning level
            if(level >= LogLevel.Warning)
                WriteMessageToTraceWindow(msg);
#endif
        }

        public static void ShowTraceWindow(bool show)
        {
            if (show)
                _traceWindow.Show();
            else
                _traceWindow.Hide();
        }

        #region Private members

        //to distinguish between user and system messages
        private enum MessageType
        {
            User,
            System
        };

        private static readonly TraceWindow _traceWindow = new TraceWindow();

        private static void WriteMessageToTraceWindow(string msg)
        {
            //TODO: check if message is identical to last one; if yes do not repeat it but append "(n times)"
            
        }

        //static constructor, ensures lazy initialisation
        //see http://csharpindepth.com/articles/general/BeforeFieldInit.aspx
        private Logger() { }

        #endregion
    }
}
