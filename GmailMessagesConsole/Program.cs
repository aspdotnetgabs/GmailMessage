using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


using System.Configuration;

namespace GmailMessagesConsole
{
    class Program
    {            //Task.Run(async () => await new GmailAPI().getEmails("INBOX"));
                 //Console.ReadKey();

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Quickstart";

        static void Main(string[] args)
        {
            var secrets = new ClientSecrets
            {
                ClientId = ConfigurationSettings.AppSettings["GMailClientId"],
                ClientSecret = ConfigurationSettings.AppSettings["GMailClientSecret"]
            };

            var token = new TokenResponse { RefreshToken = ConfigurationSettings.AppSettings["GmailRefreshToken"] };
            var credential = new UserCredential(new GoogleAuthorizationCodeFlow(
                new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = secrets
                }), "user", token);

            //Console.WriteLine("Credential file saved to: " + credPath);


            // Create Gmail API service.
            var gmailService = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });



           GetMessage(gmailService);

            // Define parameters of request.
            UsersResource.LabelsResource.ListRequest request = gmailService.Users.Labels.List("me");

            // List labels.
            IList<Label> labels = request.Execute().Labels;
            Console.WriteLine("Labels:");
            if (labels != null && labels.Count > 0)
            {
                foreach (var labelItem in labels)
                {
                    Console.WriteLine("{0}", labelItem.Name);
                }
            }
            else
            {
                Console.WriteLine("No labels found.");
            }



            Console.ReadKey();
        }

        static void GetMessage(GmailService gmailService)
        {
            var emailListRequest = gmailService.Users.Messages.List(ConfigurationSettings.AppSettings["GMailAddress"]);
            //emailListRequest.LabelIds = "INBOX";
            emailListRequest.IncludeSpamTrash = false;
            //emailListRequest.Q = "is:unread"; // This was added because I only wanted unread emails...
            emailListRequest.Q = "subject:#cocprogramming2";

            // Get our emails
            var emailListResponse = emailListRequest.Execute();

            if (emailListResponse != null && emailListResponse.Messages != null)
            {
                // Loop through each email and get what fields you want...
                foreach (var email in emailListResponse.Messages)
                {
                    var emailInfoRequest = gmailService.Users.Messages.Get(ConfigurationSettings.AppSettings["GMailAddress"], email.Id);
                    // Make another request for that email id...
                    var emailInfoResponse = emailInfoRequest.Execute();

                    if (emailInfoResponse != null)
                    {
                        String from = "";
                        String date = "";
                        String subject = "";
                        String body = "";
                        // Loop through the headers and get the fields we need...
                        foreach (var mParts in emailInfoResponse.Payload.Headers)
                        {
                            if (mParts.Name == "Date")
                            {
                                date = mParts.Value;
                            }
                            else if (mParts.Name == "From")
                            {
                                from = mParts.Value;
                            }
                            else if (mParts.Name == "Subject")
                            {
                                subject = mParts.Value;
                            }

                            if (date != "" && from != "")
                            {
                                if (emailInfoResponse.Payload.Parts == null && emailInfoResponse.Payload.Body != null)
                                {
                                    body = emailInfoResponse.Payload.Body.Data;
                                }
                                else
                                {
                                    body = getNestedParts(emailInfoResponse.Payload.Parts, "");
                                }
                                // Need to replace some characters as the data for the email's body is base64
                                //String codedBody = body.Replace("-", "+");
                                //codedBody = codedBody.Replace("_", "/");
                                //byte[] data = Convert.FromBase64String(codedBody);
                                //body = Encoding.UTF8.GetString(data);

                                // Now you have the data you want...                      
                            }

                            Console.WriteLine("{0} === {1}", from, subject);

                        }
                    }
                }
            }

        }

        static String getNestedParts(IList<MessagePart> part, string curr)
        {
            string str = curr;
            if (part == null)
            {
                return str;
            }
            else
            {
                foreach (var parts in part)
                {
                    if (parts.Parts == null)
                    {
                        if (parts.Body != null && parts.Body.Data != null)
                        {
                            str += parts.Body.Data;
                        }
                    }
                    else
                    {
                        return getNestedParts(parts.Parts, str);
                    }
                }

                return str;
            }
        }

    }



    class GmailAPI
    {
        public async Task getEmails(string labelId)
        {
            try
            {
                //UserCredential credential;
                //using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
                //{
                //    credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                //        GoogleClientSecrets.Load(stream).Secrets,
                //        // This OAuth 2.0 access scope allows for read-only access to the authenticated 
                //        // user's account, but not other types of account access.
                //        new[] { GmailService.Scope.GmailReadonly, GmailService.Scope.MailGoogleCom, GmailService.Scope.GmailModify },
                //        "NAME OF ACCOUNT NOT EMAIL ADDRESS",
                //        CancellationToken.None,
                //        new FileDataStore(this.GetType().ToString())
                //    );
                //}

                var secrets = new ClientSecrets
                {
                    ClientId = ConfigurationSettings.AppSettings["GMailClientId"],
                    ClientSecret = ConfigurationSettings.AppSettings["GMailClientSecret"]
                };

                var token = new TokenResponse { RefreshToken = ConfigurationSettings.AppSettings["GmailRefreshToken"] };
                var credential = new UserCredential(new GoogleAuthorizationCodeFlow(
                    new GoogleAuthorizationCodeFlow.Initializer
                    {
                        ClientSecrets = secrets
                    }), "user", token);

                var gmailService = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = this.GetType().ToString()
                });

                var emailListRequest = gmailService.Users.Messages.List(ConfigurationSettings.AppSettings["GMailAddress"]);
                emailListRequest.LabelIds = labelId;
                emailListRequest.IncludeSpamTrash = false;
                //emailListRequest.Q = "is:unread"; // This was added because I only wanted unread emails...

                // Get our emails
                var emailListResponse = await emailListRequest.ExecuteAsync();

                if (emailListResponse != null && emailListResponse.Messages != null)
                {
                    // Loop through each email and get what fields you want...
                    foreach (var email in emailListResponse.Messages)
                    {
                        var emailInfoRequest = gmailService.Users.Messages.Get(ConfigurationSettings.AppSettings["GMailAddress"], email.Id);
                        // Make another request for that email id...
                        var emailInfoResponse = await emailInfoRequest.ExecuteAsync();

                        if (emailInfoResponse != null)
                        {
                            String from = "";
                            String date = "";
                            String subject = "";
                            String body = "";
                            // Loop through the headers and get the fields we need...
                            foreach (var mParts in emailInfoResponse.Payload.Headers)
                            {
                                if (mParts.Name == "Date")
                                {
                                    date = mParts.Value;
                                }
                                else if (mParts.Name == "From")
                                {
                                    from = mParts.Value;
                                }
                                else if (mParts.Name == "Subject")
                                {
                                    subject = mParts.Value;
                                }

                                if (date != "" && from != "")
                                {
                                    if (emailInfoResponse.Payload.Parts == null && emailInfoResponse.Payload.Body != null)
                                    {
                                        body = emailInfoResponse.Payload.Body.Data;
                                    }
                                    else
                                    {
                                        body = getNestedParts(emailInfoResponse.Payload.Parts, "");
                                    }
                                    // Need to replace some characters as the data for the email's body is base64
                                    String codedBody = body.Replace("-", "+");
                                    codedBody = codedBody.Replace("_", "/");
                                    byte[] data = Convert.FromBase64String(codedBody);
                                    body = Encoding.UTF8.GetString(data);

                                    // Now you have the data you want...                      
                                    Console.WriteLine(from);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Failed to get messages!", "Failed Messages!");
            }
        }

        String getNestedParts(IList<MessagePart> part, string curr)
        {
            string str = curr;
            if (part == null)
            {
                return str;
            }
            else
            {
                foreach (var parts in part)
                {
                    if (parts.Parts == null)
                    {
                        if (parts.Body != null && parts.Body.Data != null)
                        {
                            str += parts.Body.Data;
                        }
                    }
                    else
                    {
                        return getNestedParts(parts.Parts, str);
                    }
                }

                return str;
            }
        }

    }
}
