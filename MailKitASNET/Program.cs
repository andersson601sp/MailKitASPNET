using System;
using System.Threading.Tasks;

namespace MailKitASPNET
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var mailRepository = new OutlookEmails("outlook.live.com", 993, true, "xxxx@hotmail.com", "xxxxx");
            var allEmails = await mailRepository.GetAllMailsAsync();

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
