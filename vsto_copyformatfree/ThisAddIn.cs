using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Clipboard = System.Windows.Forms.Clipboard;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace vsto_copyformatfree
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        public void click()
        {
            Outlook.MAPIFolder selectedFolder =
                this.Application.ActiveExplorer().CurrentFolder;

            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        string strippedBody = Regex.Replace(mailItem.Body, "<.*?>", String.Empty);
                        string copyText = string.Format("To: {0}\nFrom: {1}\nDate: {2}\nSubject: {3}\nBody: {4}",
                            mailItem.To,
                            mailItem.Sender.Address,
                            mailItem.ReceivedTime,
                            mailItem.Subject,
                            strippedBody);

                        Clipboard.SetDataObject(copyText);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Andrew error:" + ex);
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
