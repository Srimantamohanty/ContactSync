
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ContactsPrj
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            UsersResource.LabelsResource.ListRequest request = service.Users.Labels.List("me");

            //UsersResource.MessagesResource.ListRequest messageRequest = service.Users.Messages.List("ALL");

            // List labels.
            //IList<Label> labels = request.Execute().Labels;
            //Console.WriteLine("Labels:");
            //if (labels != null && labels.Count > 0)
            //{
            //    foreach (var labelItem in labels)
            //    {

            //        // Console.WriteLine("{0}", labelItem.Name);
            //        if (labelItem.Name=="INBOX")
            //        {
            //            Console.WriteLine("{0}", labelItem.Name);
            //            //bool IsSuccess = GetEmailMessageDetails(service.Users.Messages.List("me"));
            //        }
            //    }
            //}
            //else
            //{
                
            //    Console.WriteLine("No labels found.");
            //}
            
            var query = service.Users.Messages.List("me");

            var mail = query.Execute();

            foreach (var item in mail.Messages)
            {
            string From="";
            string Subject = "";
            string ReceivedDate ="";
            Message massage = service.Users.Messages.Get("me", item.Id).Execute();
            IList <MessagePartHeader> Hrds = massage.Payload.Headers;

                bool chk = false;


                foreach (var hrd in Hrds)
            {
                    chk = false;  
            IList<string> lbls = massage.LabelIds;
            foreach (var lbl in lbls)
            {
                        if (lbl == "INBOX")
                            chk = true;        
            }

            if (hrd.Name == "From")
                From = hrd.Value;
            if (hrd.Name == "Subject")
                Subject = hrd.Value;
            if (hrd.Name == "Date")
                ReceivedDate = hrd.Value;
            }

                bool inserted = false;
                if(chk)
                inserted = InsertToAZURE(From, Subject, "", ReceivedDate);
                //hello
               
            }

            Console.Read();
            
        }


        private static bool InsertToAZURE(string From, string Subject, string TextBody, string ReceivedDate)
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
                    sb.Append("INSERT INTO [dbo].[STG_Emails]([EmailFrom],[Emailheader],[EmailDetails],Date_last_interaction,Source_Of_Email)");
                    sb.Append("VALUES (@EmailFrom,@Emailheader, @EmailDetails,@Date_last_interaction,@Source_Of_Email);");
                    String sql = sb.ToString();
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@EmailFrom", From);
                        command.Parameters.AddWithValue("@Emailheader", Subject);
                        command.Parameters.AddWithValue("@EmailDetails", TextBody);
                        command.Parameters.AddWithValue("@Date_last_interaction", ReceivedDate);
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