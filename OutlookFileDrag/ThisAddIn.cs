using System;
using System.Reflection;
using log4net;

namespace OutlookFileDrag
{
    public partial class ThisAddIn
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private System.Threading.Timer cleanupTimer;
        private DragDropHook hook;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Configure logging
            log4net.Config.XmlConfigurator.Configure();

            try
            {                
                log.Info("Add-in startup");

                //Log version, OS version, Outlook version, and language
                log.InfoFormat("Version: {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
                log.InfoFormat("OS: {0} {1}", Environment.OSVersion, Environment.Is64BitOperatingSystem ? "x64" : "x86");
                log.InfoFormat("Outlook version: {0} {1}", this.Application.Version, Environment.Is64BitProcess ? "x64" : "x86");
                log.InfoFormat("Language: {0}", Application.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI));

                //Set up exception handlers
                System.Windows.Forms.Application.ThreadException += Application_ThreadException;
                AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

                //Start cleanup timer
                int cleanupTimerInterval = int.Parse(System.Configuration.ConfigurationManager.AppSettings["CleanupTimerInterval"]);
                log.InfoFormat("Starting cleanup timer -- run every {0} minutes", cleanupTimerInterval);
                cleanupTimer = new System.Threading.Timer(CleanupTimer_Callback, null, 0, cleanupTimerInterval * 60 * 1000);

                //Start hook;
                hook = new DragDropHook();
                hook.Start();
            }
            catch (Exception ex)
            {
                log.Fatal("Fatal error", ex);
                if (hook != null)
                    hook.Stop();
            }
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            log.Fatal("Appdomain exception", (Exception)e.ExceptionObject);
        }

        private void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            log.Fatal("Application thread exception", e.Exception);
        }

        private void CleanupTimer_Callback(object state)
        {
            try
            {
                int tempFileExpiration = int.Parse(System.Configuration.ConfigurationManager.AppSettings["TempFileExpiration"]);
                log.InfoFormat("Cleaning up temp files older than {0} minutes", tempFileExpiration);
                FileUtility.CleanupTempPath(tempFileExpiration);
            }
            catch (Exception ex)
            {
                log.Error("Error cleaning up temp files", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785

            try
            {
                log.Info("Add-in shutdown");
                if (hook != null)
                    hook.Stop();
            }
            catch (Exception ex)
            {
                log.Fatal("Fatal error", ex);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
