using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Timers;
using System.Xml;
using System.Xml.Serialization;
using System.Net.Mail;

namespace k163825_Q3
{
    public partial class SMTPWinSrvc : ServiceBase
    {
        private System.Timers.Timer timer;
        string emailFile = ConfigurationManager.AppSettings["xmlPath"];
        string nextEmail = ConfigurationManager.AppSettings["LastEmail"];
        int emailnum = 0;

        
        
        public SMTPWinSrvc()
        {
            InitializeComponent();
            
            timer = new System.Timers.Timer();
            timer.Interval = 15 *60 * 1000;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(WorkProcess);
        }

        private void WorkProcess(object sender, ElapsedEventArgs e)
        {
            string receiverAddr;
            string subjectMail;
            string mailMsgBody;
            XmlDocument xmlDoc = new XmlDocument();
            FileStream fs = new FileStream(nextEmail, FileMode.OpenOrCreate, FileAccess.Read);
            fs.Close();

            string[] files = Directory.GetFiles(emailFile);
            if (new FileInfo(nextEmail).Length == 0)
            {
                writeOnFile(emailnum);
            }
            emailnum = readFromFile();

            int i = 0;
            foreach(string file in files) {
                string em = "\\Email" + (emailnum+i) + ".xml";
                string email = emailFile + em;
                if (File.Exists(email))
                {
                    List<String> emailArtifacts = new List<String>();
                    Debug.WriteLine("Sending Email " + (emailnum));

                    xmlDoc.Load(email);

                    foreach (XmlNode node in xmlDoc.DocumentElement)
                    {
                        foreach (XmlNode child in node.ChildNodes)
                        {
                            emailArtifacts.Add(child.InnerText);
                        }
                    }

                    receiverAddr = emailArtifacts[0];
                    subjectMail = emailArtifacts[1];
                    mailMsgBody = emailArtifacts[2];
                    genrateEmail(receiverAddr, subjectMail, mailMsgBody);
                    emailnum = emailnum + i;
                    writeOnFile(emailnum);
                }
                i++;
            }
        }

        private void genrateEmail(string receiverAddr, string subjectMail, string mailMsgBody)
        {
            SmtpClient smtpClient = new System.Net.Mail.SmtpClient();
            smtpClient.EnableSsl = true;
            MailMessage mailMesg = new MailMessage();
            System.Net.Mime.ContentType XMLtype = new System.Net.Mime.ContentType("text/html");
            mailMesg.BodyEncoding = System.Text.Encoding.Default;
            mailMesg.To.Add(receiverAddr);
            mailMesg.Priority = System.Net.Mail.MailPriority.High;
            mailMesg.Subject = subjectMail;
            mailMesg.Body = mailMsgBody;
            mailMesg.IsBodyHtml = true;
            AlternateView alternateView = AlternateView.CreateAlternateViewFromString(mailMsgBody, XMLtype);
            smtpClient.Send(mailMesg);
        }

        public int readFromFile()
        {
            string nextemail = ConfigurationManager.AppSettings["LastEmail"];
            //Folder Must Exists
            FileStream fs = new FileStream(nextEmail, FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs);
            string s = sr.ReadLine();
            sr.Close();
            return Int32.Parse(s);
        }

        public void writeOnFile(int i)
        {
            string nextemail = ConfigurationManager.AppSettings["LastEmail"];
            FileStream fs = new FileStream(nextemail, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(i.ToString());
            sw.Flush();
            sw.Close();
        }

        protected override void OnStart(string[] args)
        {
            timer.AutoReset = true;
            timer.Enabled = true;
            LogService("Service Started");
        }

        public void OnDebug()
        {
            OnStart(null);
        }

        protected override void OnStop()
        {
            timer.AutoReset = false;
            timer.Enabled = false;
            LogService("Service Stopped");
        }

        private void LogService(string content)
        {
            string logPath = ConfigurationManager.AppSettings["LogPath"];
            //Folder Must Exists
            FileStream fs = new FileStream(logPath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }
    }
}
