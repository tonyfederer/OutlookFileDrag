using System;
using log4net;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

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

                //Start cleanup timer
                int cleanupTimerInterval = int.Parse(System.Configuration.ConfigurationManager.AppSettings["CleanupTimerInterval"]);
                log.InfoFormat("Starting cleanup timer -- run every {0} minutes", cleanupTimerInterval);
                cleanupTimer = new System.Threading.Timer(CleanupTimer_Callback, null, 0, cleanupTimerInterval * 60 * 1000);

                //Hook active explorer ViewChange event
                log.Info("Hooking explorer ViewSwitch event");
                explorer = this.Application.ActiveExplorer();
                explorer.ViewSwitch += Explorer_ViewSwitch;

                //Hook drag and drop event
                StartHook();
            }
            catch (Exception ex)
            {
                log.Fatal("Fatal error", ex);
                StopHook();
            }
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
                Outlook.View view = this.Application.ActiveExplorer().CurrentView;
                if (view.ViewType == Outlook.OlViewType.olCalendarView)
                    StopHook();
                else
                    StartHook();
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
            //Start hooking drag and drop
            if (Hook == null)
            {
                Hook = new DragDropHook();
                Hook.StartHook();
            }
        }

        private void StopHook()
        {
            //Stop hooking drag and drop
            if (Hook != null)
            {
                Hook.Dispose();
                Hook = null;
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
