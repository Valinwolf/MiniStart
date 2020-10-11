using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace MiniStart
{
    public partial class ThisAddIn
    {
        System.Timers.Timer timer = new System.Timers.Timer(500);
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Startup += Application_Startup;
            timer.Elapsed += Timer_Elapsed;
        }

        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            timer.Stop();
            timer.Elapsed -= Timer_Elapsed;
            timer.Dispose();
            Application.Startup -= Application_Startup;
            Minimize();
        }

        private void Application_Startup()
        {
            timer.Start();
        }

        private void Minimize()
        {
            Application.ActiveExplorer().WindowState = Outlook.OlWindowState.olMinimized;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
