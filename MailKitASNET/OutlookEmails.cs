using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MailKitASPNET
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }
        public List<string> PathAttachments { get; set; }

        public List<string> KeyFields { get; set; }

        private readonly string mailServer, login, password;
        private readonly int port;
        private readonly bool ssl;

        public OutlookEmails()
        {

        }

        public OutlookEmails(string mailServer, int port, bool ssl, string login, string password)
        {
            this.mailServer = mailServer;
            this.port = port;
            this.ssl = ssl;
            this.login = login;
            this.password = password;
        }

        public IEnumerable<OutlookEmails> GetAllMails()
        {
            var messages = new List<OutlookEmails>();

            using (var client = new ImapClient())
            {
                client.CheckCertificateRevocation = false;
                client.Connect(mailServer, port, ssl);
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate(login, password);

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                var results = inbox.Search(SearchQuery.Not(SearchQuery.Seen).And(SearchQuery.FromContains("exemplo@mail.com")));

                foreach (var uniqueId in results)
                {
                    var message = inbox.GetMessage(uniqueId);

                    OutlookEmails mail = new OutlookEmails
                    {
                        EmailBody = message.TextBody,
                        EmailFrom = message.From.ToString(),
                        EmailSubject = message.Subject,
                        PathAttachments = new List<string>()
                    };

                    // atttac
                    if (message.Attachments.Count() > 0)
                    {
                        var folder = $"D:\\test\\{uniqueId.Id}";
                        if (!Directory.Exists(folder))
                        {
                            Directory.CreateDirectory(folder);
                        }

                        foreach (var attachment in message.Attachments)
                        {
                            var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                            var file = $"{folder}\\{fileName}";
                            using (var stream = File.Create(file))
                            {
                                if (attachment is MessagePart)
                                {
                                    var rfc822 = (MessagePart)attachment;

                                    rfc822.Message.WriteTo(stream);
                                }
                                else
                                {
                                    var part = (MimePart)attachment;

                                    part.Content.DecodeTo(stream);
                                }
                                mail.PathAttachments.Add(file);
                            }
                        }
                    }

                    messages.Add(mail);
                    inbox.AddFlags(uniqueId, MessageFlags.Seen, true);
                }

                client.Disconnect(true);
            }

            return messages;
        }

    }
}
