using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net.Mail;
using _001TN0173.Entities;
using System.Globalization;

namespace _001TN0173.Shared
{
    class Email
    {
        ClsErrorLog clserr = new ClsErrorLog();
        Clsbase objBaseClass = new Clsbase();
        Settings sttg = new Settings();

        public void SendEmail(string ToEmail, string SendCC, string SendBCC, string AttachmentFile, string subject, string Matter)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                mail.From = new MailAddress(sttg.Email_FromID);
                smtp.Port = Convert.ToInt32(sttg.Port);
                smtp.Host = sttg.SMTP_HOST;

                string [] AryEmail = ToEmail.ToString().Split(';');
                for (int E=0; E < AryEmail.Length ;E++)
                {
                    mail.To.Add(AryEmail[E]);
                }
                
                mail.CC.Add(SendCC);
                mail.Bcc.Add(SendBCC);

                mail.Subject = subject;
                
                string strMailBody = "";
                mail.IsBodyHtml = true;
                if (AttachmentFile.ToString().Trim().Length > 0)
                {
                    System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(AttachmentFile);
                    mail.Attachments.Add(attachment);
                }

                strMailBody = "<HTML><body><br><br>";
                strMailBody = strMailBody + Matter;
                mail.Body = strMailBody;

                //AlternateView htmlView = new AlternateView();

                //htmlView = AlternateView.CreateAlternateViewFromString(strMailBody,null,"text/html");
                //mail.AlternateViews(htmlView);

                smtp.UseDefaultCredentials = true;
                smtp.Credentials = new System.Net.NetworkCredential(sttg.UserID.ToString(), sttg.Password);
                smtp.Send(mail);

                clserr.LogEntry(ToEmail + " Email Sent Successfully", false);

                DateTime dtStartDate1 = DateTime.ParseExact(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss").ToString().Replace("-", "/").Replace(".", "/"), "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture); //Current Date
                clserr.LogEntry("No invoice record found " + dtStartDate1, false);

            }
            catch (Exception ex)
            {
                clserr.LogEntry("Attached file = " + Path.GetFileName(AttachmentFile), false);
                clserr.Handle_Error(ex, "Form", "SendExceptionReport");
            }

        }

    }
}
