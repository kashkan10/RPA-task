using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Chrome;
using TaskRPA.Messenger;
using TaskRPA.Writer;

namespace TaskRPA
{
    class Program
    {
        static void Main(string[] args)
        {
            string emailValidation = @"(\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)";
            string documentPath = ConfigurationManager.AppSettings["documentPath"];
            string directoryPath = ConfigurationManager.AppSettings["directoryPath"];
            string recipientMail;

            if (args.Length != 0)
            {
                recipientMail = args.First();
            }
            else
            {
                Console.WriteLine("Enter email of recipient:");
                recipientMail = Console.ReadLine();
            }

            while (!Regex.IsMatch(recipientMail, emailValidation))
            {
                Console.WriteLine("Incorrect email, please try again:");
                recipientMail = Console.ReadLine();
            }

            try
            {
                Console.WriteLine("Start parsing...");
                Parser parser = new Parser(new ChromeDriver());
                var microwaves = parser.Parse();
                Console.WriteLine("OK");

                CreateDirectory(directoryPath);

                Console.WriteLine("Start writing...");
                var excelWriter = new ExcelWriter(documentPath);
                excelWriter.Write(microwaves);
                Console.WriteLine("OK");

                Console.WriteLine("Sending message...");
                IMessenger messenger = new OutlookMessenger(recipientMail, documentPath);
                messenger.Send();
                Console.WriteLine("OK");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Press enter to exit.");
                Console.ReadLine();
            }
        }

        private static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }
    }
}
