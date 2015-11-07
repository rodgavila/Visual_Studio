using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace GmailQuickstart
{
    class Program
    {
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Quickstart";
        static string ATTACHMENTS_PATH = @"C:\Users\Rodrigo\Documents\Biopsies";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = GetLocalCredentialsPath();

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

            Label inbox = getInboxLabel(service, "me");

            List<Message> biopsyMessages = GetAllBiopsyMessages(service, "me", inbox, "");

            foreach (Message m in biopsyMessages)
            {
                Message fullMessage = GetMessage(service, "me", m.Id);
                String subject = null;
                String from = null;
               
                foreach (MessagePartHeader header in fullMessage.Payload.Headers)
                {
                    if (header.Name.Equals("Subject", StringComparison.InvariantCultureIgnoreCase))
                    {
                        subject = header.Value;
                    } else if (header.Name.Equals("From", StringComparison.InvariantCultureIgnoreCase))
                    {
                        from = header.Value; 
                    }
                }

                Debug.Assert(subject != null && from != null, "Either subject or from are null. What happened!!");

                String biopsyId = ExtractBiopsyIdFromTitle(subject);
                
                if (biopsyId != null)
                {
                    Console.WriteLine("Processing {0}", biopsyId);
                    String biopsyType = ExtractBiopsyTypeFromTitle(subject);
                    String attachmentsPath = GetAttachmentsFolderFor(biopsyId, biopsyType);
                    Directory.CreateDirectory(attachmentsPath);

                    GetAttachments(service, "me", fullMessage.Id, attachmentsPath);
                }
            }

            Console.ReadLine();
        }

        private static string GetAttachmentsFolderFor(string biopsyId, string biopsyType)
        {
            return Path.Combine(ATTACHMENTS_PATH, biopsyId, biopsyType);
        }

        private static string GetLocalCredentialsPath()
        {
            string personalFolder = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
            return Path.Combine(personalFolder, ".credentials/gmail-dotnet-quickstart");
        }

        private static string ExtractBiopsyIdFromTitle(string title)
        {
            Regex biopsyIdRegex = new Regex(@"E-\d+-\d+", RegexOptions.IgnoreCase);
            Match match = biopsyIdRegex.Match(title);

            if (match.Success)
            {
                return match.Value;
            }

            return null;
        }

        private static string ExtractBiopsyTypeFromTitle(string title)
        {
            String biopsyId = ExtractBiopsyIdFromTitle(title);

            if (biopsyId == null)
                return null;

            int indexOfId = title.IndexOf(biopsyId);

            if (indexOfId == 0)
            {
                return title.Substring(biopsyId.Length, title.Length - biopsyId.Length).Trim();
            } else if (indexOfId == (title.Length - biopsyId.Length))
            {
                return title.Substring(0, indexOfId).Trim();
            } else
            {
                return title.Substring(0, indexOfId).Trim() + " " +
                        title.Substring(indexOfId+biopsyId.Length, title.Length - (indexOfId+biopsyId.Length)).Trim();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        private static Label getInboxLabel(GmailService service, String userId)
        {
            UsersResource.LabelsResource.ListRequest request = service.Users.Labels.List(userId);

            IList<Label> labels = request.Execute().Labels;
            Console.WriteLine("Labels:");
            if (labels != null && labels.Count > 0)
            {
                foreach (Label labelItem in labels)
                {
                    if (String.Compare(labelItem.Id, "Inbox", StringComparison.InvariantCultureIgnoreCase)==0)
                    {
                        return labelItem;
                    }            
                }
            }
            else
            {
                Debug.Assert(false, "Inbox wasn't found");
            }

            return null;
        }

        /// <summary>
        /// Retrieve a Message by ID.
        /// </summary>
        /// <param name="service">Gmail API service instance.</param>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message to retrieve.</param>
        public static Message GetMessage(GmailService service, String userId, String messageId)
        {
            try
            {
                return service.Users.Messages.Get(userId, messageId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }

            return null;
        }

        /// <summary>
        /// List all Messages of the user's mailbox matching the query.
        /// </summary>
        /// <param name="service">Gmail API service instance.</param>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="query">String used to filter Messages returned.</param>
        public static List<Message> GetAllBiopsyMessages(GmailService service, String userId, Label label, String query)
        {
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.LabelIds = label.Id;
            /// request.Q = query;

            do
            {
                try
                {
                    ListMessagesResponse response = request.Execute();
                    result.AddRange(response.Messages);
                    request.PageToken = response.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return result;
        }

        /// <summary>
        /// Get and store attachment from Message with given ID.
        /// </summary>
        /// <param name="service">Gmail API service instance.</param>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message containing attachment.</param>
        /// <param name="outputDir">Directory used to store attachments.</param>
        public static void GetAttachments(GmailService service, String userId, String messageId, String outputDir)
        {
            try
            {
                Message message = service.Users.Messages.Get(userId, messageId).Execute();
                IList<MessagePart> parts = message.Payload.Parts;
                foreach (MessagePart part in parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        String attId = part.Body.AttachmentId;
                        MessagePartBody attachPart = service.Users.Messages.Attachments.Get(userId, messageId, attId).Execute();

                        // Converting from RFC 4648 base64-encoding
                        // see http://en.wikipedia.org/wiki/Base64#Implementations_and_history
                        String attachData = attachPart.Data.Replace('-', '+');
                        attachData = attachData.Replace('_', '/');

                        byte[] data = Convert.FromBase64String(attachData);
                        File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
        }

    }
}