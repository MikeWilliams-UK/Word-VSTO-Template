using System;

namespace MyFirstTemplate
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Logger.Info($"{nameof(ThisDocument_Startup)}({FullName})");
            }
            catch (Exception exception)
            {
                Logger.Error($"{nameof(ThisDocument_Startup)}()");
                Logger.Exception(exception);
            }
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                Logger.Info($"{nameof(ThisDocument_Startup)}({FullName})");
            }
            catch (Exception exception)
            {
                Logger.Error($"{nameof(ThisDocument_Startup)}()");
                Logger.Exception(exception);
            }
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
