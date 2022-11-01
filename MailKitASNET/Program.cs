using System;

namespace MailKitASPNET
{
    class Program
    {
        static void Main(string[] args)
        {
            var mailRepository = new OutlookEmails("outlook.live.com", 993, true, "xxxx@hotmail.com", "xxxxx");
            var allEmails = mailRepository.GetAllMails();

            foreach (var mail in allEmails)
            {
                Console.WriteLine("");
                Console.WriteLine("Mail Receive from " + mail.EmailFrom);
                Console.WriteLine("Mail Subject " + mail.EmailSubject);
                Console.WriteLine("MailBody ");
                Console.WriteLine(mail.EmailBody);
                Console.WriteLine("");
                Console.WriteLine("PathAttachments ");
                foreach (var item in mail.PathAttachments)
                {
                    Console.WriteLine(item);
                }

                Console.WriteLine("");
            }
        }
    }
}
