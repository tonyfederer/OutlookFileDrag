using System;
using log4net;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookFileDrag
{
    public partial class ThisAddIn
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private Outlook.Explorer explorer;
        private System.Threading.Timer cleanupTimer;

        internal DragDropHook Hook { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Configure logging
            log4net.Config.XmlConfigurator.Configure();

            try
            {                
                log.Info("Add-in startup");
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
                Hook = new DragDropHook();

                //Hook explorer ViewChange event
                log.Info("Hooking explorer ViewSwitch event");
                explorer = this.Application.ActiveExplorer();
                explorer.ViewSwitch += Explorer_ViewSwitch;

                //Start hook if not in calendar view
                Explorer_ViewSwitch();
            }
            catch (Exception ex)
            {
                log.Fatal("Fatal error", ex);
                StopHook();
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
                log.Fatal("Fatal error", ex);
                StopHook();
            }
        }

        private void Explorer_ViewSwitch()
        {
            try
            {
                //HACK: Disable drag and drop hook when in calendar mode
                //For some reason dragging an item in calendar view throws E_NOINTERFACE exception when DoDragDrop COM function is hooked (thread issue?)
                Outlook.View view = (Outlook.View)explorer.CurrentView;
                if (view.ViewType == Outlook.OlViewType.olCalendarView)
                {
                    if (Hook.IsHooked)
                    {
                        log.Info("Calendar view detected -- stopping hook");
                        StopHook();
                    }
                }
                else
                {
                    if (!Hook.IsHooked)
                    {
                        log.Info("Non-calendar view detected -- starting hook");
                        StartHook();
                    }
                }
            }
            catch (Exception ex)
            {
                log.Fatal("Fatal error", ex);
                StopHook();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785

            try
            {
                log.Info("Add-in shutdown");
                StopHook();
            }
            catch (Exception ex)
            {
                log.Fatal("Fatal error", ex);
            }
        }

        private void StartHook()
        {
            Hook.StartHook();
        }

        private void StopHook()
        {
            if (Hook != null)
                Hook.StopHook();
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
