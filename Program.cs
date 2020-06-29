/// <summary>
/// Mail sender for TLS is developed by Alexander Startsev, 2020.
/// 
/// Mail sender for  integration with Blue Prism using Start Process.
/// 
/// The process takes JSON argument, which it then parsed according to MailInputs class.
/// The parsed data is used for sending a mail using MailKit and MimeKit.
/// 
/// The mail can be sent reading data from an input text file,
/// however, the latter option is deprecated at the moment.
/// </summary>
/// 

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

using MimeKit;
using MailKit.Net.Smtp;
using MailKit.Security;
//using System.Security.Authentication;

using Newtonsoft.Json;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Crypto;

namespace MailMime_test
{



    public class MailText
    {
        public IList<MailInputs> Roles { get; set; }
    }

    public class MailInputs
    {
        public string To { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        public string Attachment { get; set; }
        public string Subject {get; set;}
        public string Body {get; set;}
        public string User { get; set; }
        public string Password { get; set; }
        public string Server { get; set; }
        public string Port { get; set; }

    }



    class Program
    {
        static async Task<string> ReadAndDisplayFilesAsync()
        {
            String filename = "C:\\Users\\ATOM\\Desktop\\packages\\SEND_MAIL.txt";
            Char[] buffer;

            using (var sr = new StreamReader(filename))
            {
                buffer = new Char[(int)sr.BaseStream.Length];
                await sr.ReadAsync(buffer, 0, (int)sr.BaseStream.Length);
            }

            string lines = new String(buffer);
            Console.WriteLine(lines);
            return lines;
        }

        static MailInputs DeserializeJson(string jsontext)
        {
            Console.WriteLine("De-serializing...");
            Console.WriteLine(jsontext.Length.ToString());
            MailInputs maildata = JsonConvert.DeserializeObject<MailInputs>(jsontext.Replace("\"","\'"));
            
                       
            Console.WriteLine("De-serialization completed: "+ maildata.To);

            return maildata;
        }

        static void SendMailMime(MailInputs maildata)
        {

            try
            {
                //---------------------
                //  Configuration

                //  User setup
                string User = maildata.User; 
                string UserName = maildata.User;        
                string pssw = maildata.Password; 

                //  Server setup
                string server = maildata.Server;  
                int port =  Convert.ToInt32(maildata.Port); 

                //  Mail setup
                string attachment = maildata.Attachment;
                string MailBody = maildata.Body;  
                string MailSubject = maildata.Subject; 

                string toAddressString = maildata.To; 
                string ccAddressString = maildata.Cc; 
                string bccAddressString = maildata.Bcc; 

                //--------------------

                MimeMessage message = new MimeMessage();

                MailboxAddress from = new MailboxAddress(UserName, User);
                message.From.Add(from);

                if (!string.IsNullOrEmpty(toAddressString))
                {
                    InternetAddressList toList = new InternetAddressList();
                    toList.AddRange(MimeKit.InternetAddressList.Parse(toAddressString));
                    message.To.AddRange(toList);   
                }


                if (!string.IsNullOrEmpty(ccAddressString))
                {
                    InternetAddressList ccList = new InternetAddressList();
                    ccList.AddRange(MimeKit.InternetAddressList.Parse(ccAddressString));
                    message.Cc.AddRange(ccList);   
                }


                if (!string.IsNullOrEmpty(bccAddressString))
                {
                    InternetAddressList bccList = new InternetAddressList();
                    bccList.AddRange(MimeKit.InternetAddressList.Parse(bccAddressString));
                    message.Bcc.AddRange(bccList);
                 }

                //Mail body

                BodyBuilder bodyBuilder = new BodyBuilder();
                //bodyBuilder.HtmlBody = "<h1>This is a mail body</h1>";
                message.Subject = MailSubject;
                if (!string.IsNullOrEmpty(MailBody))
                {
                    bodyBuilder.TextBody = MailBody;
                }


                //
                if (!string.IsNullOrEmpty(attachment))
                {
                    bodyBuilder.Attachments.Add(attachment);
                }

                message.Body = bodyBuilder.ToMessageBody();


                SmtpClient client = new SmtpClient();


                client.Connect(server, port, SecureSocketOptions.None);
                client.Authenticate(User, pssw);


                client.Send(message);
                client.Disconnect(true);
                client.Dispose();



                Console.WriteLine("");
                //Console.ReadLine();

                //............................
            }
            catch (Exception ep)
            {
                Console.WriteLine("failed to send email with the following error:");
                Console.WriteLine(ep.Message);
                Console.ReadLine();
            }

        }

        static async Task FileBasedMail()
        {
            string text = await ReadAndDisplayFilesAsync();
            Console.WriteLine("Output: "+text);
            MailInputs MailData =  Program.DeserializeJson(text);
            Program.SendMailMime(MailData);

            //DeserializeJsonTest();

        }

        static void ArgumentBasedMail(string jsoninputs)
        {
            
            //Console.WriteLine("Output: " + jsoninputs);
            MailInputs MailData = Program.DeserializeJson(jsoninputs);
            Program.SendMailMime(MailData);  

        }

        static void Sender(string jsoninputs)
        {
            
            Console.WriteLine(jsoninputs);

            if (!string.IsNullOrEmpty(jsoninputs))
            {

                Console.WriteLine("Argument is present");
                ArgumentBasedMail(jsoninputs);
                //Console.ReadLine();

            }
            else
            {

                Console.WriteLine("******There could have been FileBasedMail*********");
                //await FileBasedMail();
                Console.WriteLine("***************************************************");
                //Console.ReadLine();

            }

        }

        public static void Main(string[] args)
        {
            Console.WriteLine();
            //  Invoke this sample with an arbitrary set of command line arguments.
            Console.WriteLine("CommandLine argument: {0}", Environment.CommandLine);
            Console.WriteLine();
            
 
                int i = 0;

            //
            //Console.Out.WriteLine("******argument received*******");
            //Console.Out.WriteLine();
            //Sender(jstext);
            //Console.Out.WriteLine("*******mail sent**************");
            //

            Console.WriteLine(args.Length);
            if (args.Length != 0)
            {
                if (!string.IsNullOrEmpty(args[i]))

                {
                    Console.Out.WriteLine("******argument received*******");
                    string yo = string.Join(" ", args).Replace(",,", ",");
                    Console.Out.WriteLine("Correct CommandLine argument: " + yo);
                    Console.Out.WriteLine();
                    Sender(yo);
                    Console.Out.WriteLine("*******mail sent**************");
                    //Console.ReadLine();
                }
                else
                {
                    //Console.Out.WriteLine("xxxxx*argument_ not received*xxxxx");
                    Console.Error.WriteLine("Error E0002 reading an argument.");
                    Console.ReadLine();
                }
            }
            else
            {
                //Send mail can be invoked here using data read from an input file.
                //Sender(args[0]).Wait();

                //Console.Out.WriteLine("!!!!!xxxxx*argument_ not received*xxxxx!!!!!");
                Console.Error.WriteLine("Error E0001 argument is empty");
                Console.ReadLine();
            }


            
                //Console.ReadLine();
        }



    }
}