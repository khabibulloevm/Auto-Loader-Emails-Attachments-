using EmailReseiver.MailServices;
using System;
using System.Threading.Tasks;

namespace DataProvider
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var mailReceiver = new MailReceiverService();
            await mailReceiver.DoReceiveMail();

            Console.WriteLine("...");
            Console.Read();
        }
 
    }
}
