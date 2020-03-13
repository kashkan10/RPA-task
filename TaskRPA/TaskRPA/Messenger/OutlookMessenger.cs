using System;
using Microsoft.Office.Interop.Outlook;

namespace TaskRPA.Messenger
{
    /// <summary>
    /// Class that allows to send message with attachment to email.
    /// </summary>
    class OutlookMessenger : IMessenger
    {
        private Application outLookApp;
        private MailItem mailItem;

        public string RecipientMail { get; set; }
        public string Subject { get; set; }
        public string Message { get; set; }
        public string AttachmentPath { get; set; }

        public OutlookMessenger(string recipientMail, string attachmentPath)
        {
            AttachmentPath = attachmentPath;
            RecipientMail = recipientMail;
            Subject = "RPA Task";
            Message = "Task completed";
        }

        public OutlookMessenger(string recipientMail, string subject, string message, string attachmentPath)
        {
            AttachmentPath = attachmentPath;
            RecipientMail = recipientMail;
            Subject = subject;
            Message = message;
        }

        public void Send()
        {
            try
            {
                outLookApp = new Application();
                mailItem = (MailItem)outLookApp.CreateItem(OlItemType.olMailItem);
                mailItem.Subject = Subject;
                mailItem.Body = Message;
                mailItem.Recipients.Add(RecipientMail);
                mailItem.Attachments.Add(AttachmentPath, OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                mailItem.Send();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Unable to send message.");
                throw;
            }
            finally
            {
                outLookApp.Quit();
            }
        }
    }
}
