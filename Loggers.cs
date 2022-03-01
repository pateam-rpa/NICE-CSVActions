using log4net;
using System.Reflection;

namespace Direct.CSV.Library
{
    public static class Loggers
    {
        private static readonly ILog _log = LogManager.GetLogger("LibraryObjects");
        public static void LogDebug(string methodName, string message)
        {
            string assemblyTitle = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
            if (_log.IsDebugEnabled)
            {
                _log.Debug(string.Format("{0}: {1} - {2}", assemblyTitle, methodName, message));
            }
        }

        public static void LogError(string methodName, string message)
        {
            string assemblyTitle = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
            if (_log.IsErrorEnabled)
            {
                _log.Error(string.Format("{0}: {1} - {2}", methodName, assemblyTitle, message));
            }
        }
    }
}
