using System.Net;
using System.Net.Mail;

namespace TopluMailSed.Models
{
    public class SendMail
    {
        public void Microsoft(int SmtpPort, string SmtpHost, string GondericiAdSoyad, string GondericiMail, string GondericiPass, string AliciMail, string Baslik, string icerik)
        {
            SmtpClient sc = new SmtpClient();
            sc.Port = SmtpPort;
            sc.Host = SmtpHost;
            sc.EnableSsl = true;
            sc.Credentials = new NetworkCredential(GondericiMail, GondericiPass);

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(GondericiMail, GondericiAdSoyad);
            mail.To.Add(AliciMail);
            mail.Subject = Baslik;
            mail.IsBodyHtml = true;
            mail.Body = icerik;

            sc.Send(mail);
        }
    }
}
