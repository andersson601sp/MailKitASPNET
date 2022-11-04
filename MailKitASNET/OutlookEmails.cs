using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        private readonly string strSubject = "Requerimento Jira";

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

        public async Task<IEnumerable<OutlookEmails>> GetAllMailsAsync()
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

                // busco na pasta raiz onde mensagem não lida e com Subject Requerimento Jira
                var results = inbox.Search(SearchQuery.Not(SearchQuery.Seen).And(SearchQuery.SubjectContains(strSubject)));

                // Defino Subpasta a ser buscada no caso subFolder
                IMailFolder subFolder = null;
                if (client.Capabilities.HasFlag(ImapCapabilities.SpecialUse))
                    subFolder = client.GetFolder(SpecialFolder.Sent);

                if (subFolder == null)
                {
                    // recupera pasta raiz padrão
                    var personal = client.GetFolder(client.PersonalNamespaces[0]);

                    //  recupera subpasta subFolder
                    subFolder = await personal.GetSubfolderAsync("subfolder").ConfigureAwait(false);

                }

                // Copia da inbox para subpasta subFolder
                if (results.Count > 0)
                    client.Inbox.MoveTo(results, subFolder);

                // Abre subpasta subFolder
                subFolder.Open(FolderAccess.ReadWrite);

                // busco na pasta subFolder onde mensagem não lida e com Subject Requerimento Jira
                var newResults = subFolder.Search(SearchQuery.Not(SearchQuery.Seen).And(SearchQuery.SubjectContains(strSubject)));
                // fim subpasta

                ProcessMessage(messages, subFolder, newResults);

                client.Disconnect(true);
            }

            return messages;
        }

        private static void ProcessMessage(List<OutlookEmails> messages, IMailFolder subFolder, IList<UniqueId> newResults)
        {
            foreach (var uniqueId in newResults)
            {
                var message = subFolder.GetMessage(uniqueId);

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
                    var folder = $"D:\\test\\subfolder\\{uniqueId.Id}";
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
                subFolder.AddFlags(uniqueId, MessageFlags.Seen, true);
            }
        }
    }
}
