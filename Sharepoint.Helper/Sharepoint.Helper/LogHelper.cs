using System;
using Microsoft.SharePoint;
using System.IO;
using System.Web.Hosting;

namespace Sharepoint.Helper
{
    public class LogHelper
    {
        LogHelper() { }
        class Nested
        {
            static Nested()
            {
            }
            internal static readonly LogHelper instance = new LogHelper();
        }

        public static LogHelper Instance
        {
            get
            {
                return Nested.instance;
            }
        }

        public void Log(string msg, string description, LogSeverity severity)
        {
            bool IsOnline = false;

            String _severity = String.Empty;

            switch (severity)
            {
                case LogSeverity.Debug:
                    _severity = "Debug";
                    break;
                case LogSeverity.Warning:
                    _severity = "Warning";
                    break;
                case LogSeverity.Error:
                    _severity = "Error";
                    break;
                default:
                    break;
            }

            try
            {
                //if (SPContext.Current != null)
                //{
                //    SPSite rootSite = SPContext.Current.Site;
                //    bool unsafeUpdateEnabled = SPContext.Current.Site.RootWeb.AllowUnsafeUpdates;
                //    if (msg.Length > 250) { msg = msg.Substring(0, 250); }
                //    SPList logList = rootSite.RootWeb.Lists.TryGetList("Logs");
                //    if (logList != null)
                //    {
                //        IsOnline = true;
                //        rootSite.RootWeb.AllowUnsafeUpdates = true;
                //        SPListItem newLog = logList.AddItem();
                //        newLog["Title"] = msg;
                //        newLog["Description"] = description;
                //        newLog["Severity"] = _severity;
                //        newLog.Update();
                //        if (unsafeUpdateEnabled == false) { rootSite.RootWeb.AllowUnsafeUpdates = false; };
                //    }
                //    rootSite.Dispose();
                //}

                if (!IsOnline)
                {
                    StreamWriter sw = File.AppendText(HostingEnvironment.MapPath("~") + "log.txt");
                    try
                    {
                        string logLine = String.Format("'{0}';'{1}';'{2}'", _severity , msg, description);
                        sw.WriteLine(logLine);
                    }
                    finally
                    {
                        sw.Close();
                    }
                }

                System.Diagnostics.Debug.WriteLine(String.Format("'{0}';'{1}';'{2}'", _severity, msg, description));

            }
            catch (Exception)
            { }

        }

        public void Log(string msg)
        {
            Log(msg, LogSeverity.Debug);
        }

        public void Log(string msg,  LogSeverity severity)
        {
            Log(msg, String.Empty, severity);
        }

    }

    public enum LogSeverity
    {
        Debug = 0,
        Warning = 1,
        Error = 2
    }
}


