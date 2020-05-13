using System;
using System.Text.RegularExpressions;
using Clipboard = System.Windows.Forms.Clipboard;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace vsto_copyformatfree
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        public static string filterBody(string text)
        {
            return Regex.Replace(text, "(\r?\n){2,}", "\r\n");
        }

        public void click()
        {
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        string copyText = string.Format("To: {0}\nFrom: {1}\nDate: {2}\nSubject: {3}\nBody: {4}",
                            mailItem.To,
                            mailItem.Sender.Address,
                            mailItem.ReceivedTime,
                            mailItem.Subject,
                            ThisAddIn.filterBody(mailItem.Body));

                        Clipboard.SetDataObject(copyText);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log errors ?
            }
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
