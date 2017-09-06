using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AsposeOutlookExample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("enter password");
            string password = Console.ReadLine();

            SendMail(password, false);
            SendMail(password, true);

            Console.ReadKey();

        }

        private static void SendMail(string password, bool includeTimezone)
        {
            string fromEmail = "joncoelloaccess@gmail.com";
            string toEmail = "jon.coello@theaccessgroup.com";

            // client
            var client = new SmtpClient("smtp.gmail.com", 587, fromEmail, password);

            // message
            var message = new MailMessage(fromEmail, toEmail);
            message.Subject = "invite";
            message.Body = "here's your invite";

            // appointment
            var currentDate = DateTime.Now;
            var attendees = new MailAddressCollection();
            attendees.Add(toEmail);

            var startTime = new DateTime(2017,9,6,9,0,0);
            var endTime = startTime.AddHours(1);

            var appt = new Appointment("London", startTime, endTime, new MailAddress(fromEmail), attendees);

            //appt.Save(@"C:\scratch\Training Event Wizard1.ics", AppointmentSaveFormat.Ics);

            if (includeTimezone)
            {
                var tz = TimeZone.CurrentTimeZone;
                appt.SetTimeZone(tz.StandardName);
            }

            appt.Save(@"C:\scratch\Training Event Wizard3.ics", AppointmentSaveFormat.Ics);

            //message.Attachments.Add(new Attachment(@"C:\scratch\Training Event Wizard2.ics", string.Format("Calendar{0}.ics", 1)));
            message.Attachments.Add(new Attachment(@"C:\scratch\Training Event Wizard3.ics", "text/calendar"));

            client.Send(message);

            message.Dispose();

            Console.WriteLine("done");
        }
    }
}
