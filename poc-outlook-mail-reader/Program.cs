using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace poc_outlook_mail_reader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting email fetch");

            // Connect to the exchange server
            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            exchangeService.Credentials = new NetworkCredential(ConfigHelper.EmailAddress, ConfigHelper.EmailPassword);
            exchangeService.Url = new Uri(ConfigHelper.ExchangeUrl);
            exchangeService.KeepAlive = true;

            // Fetch inbox mails
            Mailbox mailbox = new Mailbox(ConfigHelper.EmailAddress, "");
            FolderId folder = new FolderId(WellKnownFolderName.Inbox, mailbox);
            ItemView view = new ItemView(999);
            FindItemsResults<Item> items = exchangeService.FindItems(folder, view);

            foreach (var i in items)
            {
                // Load message details
                i.Load();

                //Then use the message data as needed
                var messageBody = i.Body;
            }
        }
    }
}
