using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using System.Globalization;
using MimeKit.Text;
using System.IO;
using System.Net.Mime;
using System.Net.Mail;

namespace MailApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string header = $"header-of-letter";
            MimeMessage message = new MimeMessage(); 
            message.From.Add(new MailboxAddress("your-name", "dmitrijsidorchuk@yandex.ru"));
            message.To.Add(MailboxAddress.Parse("recipient-mail"));

            message.Subject = "TEST";
            string messageText = "text"; 
            TextPart body = new TextPart("plain") { Text = messageText }; 
            Multipart multipart = new Multipart("mixed");
            multipart.Add(body);

            MimePart attachment = new MimePart("application", "vnd.ms-excel")
            {
                Content = new MimeContent(File.OpenRead("file-path"), ContentEncoding.Default),
                ContentDisposition = new MimeKit.ContentDisposition(MimeKit.ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = Path.GetFileName("file-path")
            };

            //замена кодировки чтобы имя вложенного в письмо файла читалось корректно
            foreach (Parameter parameter in attachment.ContentType.Parameters) parameter.EncodingMethod = ParameterEncodingMethod.Rfc2047;
            foreach (Parameter parameter in attachment.ContentDisposition.Parameters) parameter.EncodingMethod = ParameterEncodingMethod.Rfc2047;

            multipart.Add(attachment);
            message.Body = multipart;

            MailKit.Net.Smtp.SmtpClient client = new MailKit.Net.Smtp.SmtpClient();
            try
            {
                client.Connect("smtp.yandex.ru", 465, true);
                client.Authenticate("your-mail", "the password issued by mail to your application");
                client.Send(message);
            }
            catch (Exception ex) 
            {
                    Console.WriteLine(ex.Message);
            }
            finally
            {
                client.Disconnect(true);
                client.Dispose();
            }
            
        }
    }
}