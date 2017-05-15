using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using EAGetMail; //add EAGetMail namespace
using System.Data.SqlClient;

namespace ContactsPrj
{
    class Program1
    {
        static void Main(string[] args)
        {

            string curpath = Directory.GetCurrentDirectory();
            string mailbox = String.Format("{0}\\inbox", curpath);

            // If the folder is not existed, create it.
            if (!Directory.Exists(mailbox))
            {
                Directory.CreateDirectory(mailbox);
            }

            // Gmail IMAP4 server is "imap.gmail.com"
            MailServer oServer = new MailServer("imap.gmail.com",
                        "singhsinghyadav123", "snehsinghyadav", ServerProtocol.Imap4);
            MailClient oClient = new MailClient("TryIt");

            // Set SSL connection,
            oServer.SSLConnection = true;

            // Set 993 IMAP4 port
            oServer.Port = 993;

            try
            {
                oClient.Connect(oServer);
                MailInfo[] infos = oClient.GetMailInfos();
                for (int i = 0; i < infos.Length; i++)
                {
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);

                    // Download email from GMail IMAP4 server
                    Mail oMail = oClient.GetMail(info);

                    Console.WriteLine("From: {0}", oMail.From.ToString());
                    Console.WriteLine("Subject: {0}\r\n", oMail.Subject);

                    // Generate an email file name based on date time.
                    System.DateTime d = System.DateTime.Now;
                    System.Globalization.CultureInfo cur = new
                        System.Globalization.CultureInfo("en-US");
                    string sdate = d.ToString("yyyyMMddHHmmss", cur);
                    string fileName = String.Format("{0}\\{1}{2}{3}.eml",
                        mailbox, sdate, d.Millisecond.ToString("d3"), i);

                    // Save email to local disk
                   // oMail.SaveAs(fileName, true);

                    bool inserted=false;
                    inserted = InsertToAZURE(oMail);


                    // Mark email as deleted in GMail account.
                   // oClient.Delete(info);
                }

                // Quit and purge emails marked as deleted from Gmail IMAP4 server.
                oClient.Quit();
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }
        }

        private static bool InsertToAZURE(Mail oMail)
        {
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "contactsql.database.windows.net";
                builder.UserID = "ContactUser";
                builder.Password = "Mobile123";
                builder.InitialCatalog = "Contact";

                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    Console.WriteLine("\nInsert data example:");
                    Console.WriteLine("=========================================\n");

                    connection.Open();
                    StringBuilder sb = new StringBuilder();
                    sb.Append("INSERT INTO [dbo].[STGEmails]([EmailFrom],[Emailheader],[EmailDetails],Date_last_interaction,Source_Of_Email)");
                    sb.Append("VALUES (@EmailFrom,@Emailheader, @EmailDetails,@Date_last_interaction,@Source_Of_Email);");
                    String sql = sb.ToString();
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@EmailFrom", oMail.From.ToString());
                        command.Parameters.AddWithValue("@Emailheader", oMail.Subject.Replace("(Trial Version)",""));
                        command.Parameters.AddWithValue("@EmailDetails", oMail.ReceivedDate + oMail.TextBody);
                        command.Parameters.AddWithValue("@Date_last_interaction", oMail.ReceivedDate);
                        command.Parameters.AddWithValue("@Source_Of_Email", "Gmail");
                        int rowsAffected = command.ExecuteNonQuery();
                        Console.WriteLine(rowsAffected + " row(s) inserted");
                    }
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }
            return true;
        }

        
    }
}
