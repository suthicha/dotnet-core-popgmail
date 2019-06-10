using System;
using System.IO;

using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;

using Newtonsoft.Json;

namespace dotnet_core_popgmail
{
    public class MailClient {
        private MailSettings _settings;
        public MailClient(MailSettings settings){
            _settings = settings;
        }

        public void Read(String subject = ""){
            if (_settings == null) return;

            using (var client = new ImapClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                client.Connect(_settings.Host, _settings.Port, _settings.isSSL);
                client.Authenticate(_settings.EmailAddress, _settings.EmailPassword);

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                // let's search for all messages received after Jan 12, 2013 with "MailKit" in the subject...
                DateTime today = DateTime.Now;
                DateTime scopeDate = today.AddDays(-5);
                var query = SearchQuery.DeliveredAfter (scopeDate)
                    .And (SearchQuery.SubjectContains (subject));

                // And (SearchQuery.Seen)                
                
                foreach (var uid in inbox.Search (query)) {
                    var message = inbox.GetMessage (uid);

                    foreach (MimeEntity attachment in message.Attachments) {
                        var fileName = Path.Combine(Utils.InitPath("temp"), attachment.ContentType.Name);

                        if (File.Exists(fileName)) File.Delete(fileName);
                        Console.WriteLine(fileName);

                        using (var stream = File.Create (fileName)) {
                            if (attachment is MessagePart) {
                                var rfc822 = (MessagePart) attachment;

                                rfc822.Message.WriteTo (stream);
                            } else {
                                var part = (MimePart) attachment;

                                part.Content.DecodeTo (stream);
                            }
                        }

                    }

                    // Delete Message from inbox.
                    inbox.AddFlags(uid, MessageFlags.Deleted, true);
                      
                }
                
                inbox.Expunge();
                client.Disconnect(true);

            }
        }

    }

}