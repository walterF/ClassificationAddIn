using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace ClassificationAddIn
{
    public partial class ThisAddIn
    {
        public Outlook.Application OutlookApplication;
        public Outlook.Inspectors OutlookInspectors;
        public Outlook.Inspector OutlookInspector;
        public Outlook.MailItem OutlookMailItem;
        public const String CONFIDENTIALSUBJECT = "[Classification: Confidential] ";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void ThisAddIn_ItemSend(object item, ref bool Cancel)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)item;
            Outlook.Recipients recipients = mailItem.Recipients;
            List<Outlook.Recipient> removeList = new List<Outlook.Recipient>();
            //String intranetDomain = mailItem.Sender.Address;
            String internalDomain = "/o=";
            foreach (Outlook.Recipient recipient in recipients) {
                //if (GetEmailAddressByRecipent(recipient).Contains(internalDomain)) {
                if (!recipient.Address.Contains(internalDomain)) {
                    removeList.Add(recipient);
                }
            }
            if (removeList.Count > 0) {
                String s = "Following external user found in email:";
                foreach (Outlook.Recipient recipient in removeList) {
                    s += "\n" + recipient.Name;
                }
                s += "\nDo you want to continus?";
                DialogResult result1 = MessageBox.Show(s, "External Sender Alert", MessageBoxButtons.YesNoCancel);
                if (result1 == DialogResult.Yes) {
                    if (mailItem.Permission == Outlook.OlPermission.olPermissionTemplate) {
                        mailItem.Subject = CONFIDENTIALSUBJECT + mailItem.Subject;
                    }
                    mailItem.Send();
                } else {
                    Cancel = true;
                }
            } else {
                if (!(mailItem.Subject.Contains("RE:") || mailItem.Subject.Contains("FW:"))) {
                    mailItem.Subject = CONFIDENTIALSUBJECT + mailItem.Subject;
                }
            }
        }

        private String GetEmailAddressByRecipent(Outlook.Recipient recipient) {
            switch (recipient.AddressEntry.AddressEntryUserType) {
                case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry: case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry:
                    return recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                case Outlook.OlAddressEntryUserType.olOutlookContactAddressEntry: case Outlook.OlAddressEntryUserType.olSmtpAddressEntry:
                    return recipient.AddressEntry.Address;
                default:
                    return recipient.Address;
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
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(ThisAddIn_ItemSend);
        }
        
        #endregion
    }

    public class StopProcessException : Exception {
        public StopProcessException() {
        }
    }
}
