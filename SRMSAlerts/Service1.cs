using SRMSAlerts.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.SqlServer;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SRMSAlerts
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer(); 

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("[LOG] Service is started at " + DateTime.Now + "\n");
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 5000; 
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteToFile("[LOG] Service is stopped at " + DateTime.Now + "\n");
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("[LOG] Check qualifications " + DateTime.Now + "\n");

            CheckQualifications();
        }

        private void CheckQualifications()
        {
            using (var context = new EAM_NL_TESTEntities1())
            {

                //Get list of records <= 30 days
                var records = context.R5PERSONNEL.Join(
                  context.R5PERSONNELQUALIFICATIONS,
                  per => per.PER_CODE,
                  per_qua => per_qua.PQU_PERSON,
                  (per, per_qua) => new {
                      CODE = per.PER_CODE,
                      DESC = per.PER_DESC,
                      TRADE = per.PER_TRADE,
                      ORG = per.PER_ORG,
                      QUALIFICATION = per_qua.PQU_QUALIFICATION,
                      START = per_qua.PQU_QUALIFICATIONSTART,
                      EXPIRATION = per_qua.PQU_EXPIRATION,
                      SQLIDENTITY = per_qua.PQU_SQLIDENTITY
                  }
                ).Where(resultTable => SqlFunctions.DateDiff("day", DateTime.Now, resultTable.EXPIRATION) <= 30)
                .ToList();


                //create list of records to be sent by email and list of saved ids
                StringBuilder emailBody = new StringBuilder();
                StringBuilder sentIDs = new StringBuilder();
                foreach (var r in records)
                {
                    string code = r.CODE;
                    if (!NotificationSent(r.SQLIDENTITY))
                    {
                        emailBody.Append(r.CODE + " " + r.DESC + " " + r.ORG + " Qualification: " + r.QUALIFICATION + " Start date: " + r.START + "End date " + r.EXPIRATION + "SqlID: " + r.SQLIDENTITY + "\n");
                        sentIDs.Append(r.SQLIDENTITY + "\n");
                    }
                }

                if(!String.IsNullOrEmpty(emailBody.ToString()))
                    SendEmail(emailBody.ToString());

                if (!String.IsNullOrEmpty(sentIDs.ToString()))
                    RegisterNotification(sentIDs.ToString());
            }
        }

        //check int LOG file if record was already sent
        private bool NotificationSent(int SQLIDENTITY)
        {
            string line = "";
            bool exists = false;

            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog.txt";
            System.IO.StreamReader file = new System.IO.StreamReader(filepath);
            while ((line = file.ReadLine()) != null)
            {
                if (line.Equals(SQLIDENTITY.ToString())) {
                    exists = true;
                    break;
                }
            }

            file.Close();

            return exists;
        }

        //send email to manager
        private void SendEmail(string body)
        {
            WriteToFile("[LOG] Email should be sent: " + DateTime.Now + "\n" + body);
            try { 
                 // add from,to mailaddresses
                 MailAddress from = new MailAddress("stadlerxrail@gmail.com", "TestFromName");
                 MailAddress to = new MailAddress("jan.godlewski@stadlerrail.com", "TestToName");
                 MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

                 // set subject and encoding
                 myMail.Subject = "Test message";
                 myMail.SubjectEncoding = System.Text.Encoding.UTF8;

                 // set body-message and encoding
                 myMail.Body = "<b>Test Mail</b><br>" + "<b>HTML</b>.";
                 myMail.BodyEncoding = System.Text.Encoding.UTF8;
                 myMail.IsBodyHtml = true;

                 SmtpClient mySmtpClient = new SmtpClient("dmzrelay.stadlerrail.ch", 587);
                 mySmtpClient.UseDefaultCredentials = false;
                 mySmtpClient.Credentials = new NetworkCredential("stadlerxrail@gmail.com", "Stadler1234");
                 mySmtpClient.EnableSsl = true;
                 mySmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                 mySmtpClient.Send(myMail);

             }
             catch (SmtpException ex)
             {
                 throw new ApplicationException
                   ("SmtpException has occured: " + ex.Message);
             }
             catch (Exception ex)
             {
                 throw ex;
             }
        }

        //Log that ID was sent in the LOG file
        private void RegisterNotification(string listID)
        {
            WriteToFile("[LOG] ID were sent: " + DateTime.Now + "\n" + listID);
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog.txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.Write(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.Write(Message);
                }
            }
        }
    }
}
