using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace ConsoleApp2
{
    class Program
    {
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "Gmail attachment downloader";
        static List<string> fileTypes = new List<string> { ".png", ".jpg", ".jpeg", ".bmp" }; //used for deletions
        static int interval = 20; //time between checks in seconds
        static string downloadPath = @"C:\Users\jisacd1\Desktop\Test"; //put attachments here
        static int fileSizeLowerBound = 100000; //don't pull if it's smaller than this

        static void Main(string[] args)
        {
            UserCredential credential = GetGmailCredentials();
            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            long? lastSeenDate = 1;
            int uniqueNum = 0;

            DeleteFiles(downloadPath, fileTypes);

            while (true)
            {
                //Console.WriteLine($"DEBUG: Current last seen date: {lastSeenDate}");

                List<string> messageIDs = GetMessages(service, ref lastSeenDate);
                
                Console.WriteLine($"Found {messageIDs.Count} new messages from Gmail.");

                foreach (string ID in messageIDs)
                {
                    Console.WriteLine($"Pulling attachments...");

                    int count = GetAttachments(service, "me", ID, downloadPath, uniqueNum);

                    Console.WriteLine($"Put {count} attachments into {downloadPath}.");
                }

                Console.WriteLine($"Waiting {interval} seconds.");

                System.Threading.Thread.Sleep(1000 * interval);
            }
        }

        #region Helpers

        private static void DeleteFiles(string downloadPath, List<string> filetypes)
        {
            foreach (string filePath in Directory.GetFiles(downloadPath))
            {
                string fileType = Path.GetExtension(filePath);

                if (fileTypes.Contains(fileType))
                {
                    File.Delete(filePath);
                }
            }
        }

        private static UserCredential GetGmailCredentials()
        {
            UserCredential credential;

            using (var stream =
             new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }

        private static List<string> GetMessages(GmailService service, ref long? lastSeenDate)
        {
            var getmessages = service.Users.Messages.List("me");
            long? biggestDate = lastSeenDate;
            IList<Message> thinMessageList = getmessages.Execute().Messages;
            List<Message> messages = new List<Message>();

            foreach(Message message in thinMessageList)
            {
                var fullMessageRequest = service.Users.Messages.Get("me", message.Id);
                messages.Add(fullMessageRequest.Execute());
            }

            List<string> messageIDs = new List<string>();

            //Console.WriteLine($"DEBUG: {messages.Count} messages found in the inbox.");

            foreach (Message message in messages)
            {
                long? thisDate = message.InternalDate;

                //Console.WriteLine($"DEBUG: {thisDate} > {lastSeenDate} ?");

                //if message is new, add it to our list
                if (thisDate > lastSeenDate)
                {
                    string messageID = message.Id;
                    messageIDs.Add(messageID);

                    //keep track of the biggest date, without updating last seen yet
                    if (thisDate > biggestDate)
                    {
                        biggestDate = message.InternalDate;
                    }
                }
            }
            //now update last seen
            lastSeenDate = biggestDate;
            return messageIDs;
        }

        /// <summary>
        /// Get and store attachment from Message with given ID.
        /// </summary>
        /// <param name="service">Gmail API service instance.</param>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message containing attachment.</param>
        /// <param name="outputDir">Directory used to store attachments.</param>
        public static int GetAttachments(GmailService service, string userId, string messageId, string outputDir, int uniqueNum)
        {
            int count = 0;
            try
            {
                Message message = service.Users.Messages.Get(userId, messageId).Execute();
                IList<MessagePart> parts = message.Payload.Parts;
                foreach (MessagePart part in parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename) && part.Body.Size > fileSizeLowerBound)
                    {
                        count++;
                        String attId = part.Body.AttachmentId;
                        MessagePartBody attachPart = service.Users.Messages.Attachments.Get(userId, messageId, attId).Execute();

                        // Converting from RFC 4648 base64 to base64url encoding
                        // see http://en.wikipedia.org/wiki/Base64#Implementations_and_history
                        String attachData = attachPart.Data.Replace('-', '+');
                        attachData = attachData.Replace('_', '/');

                        byte[] data = Convert.FromBase64String(attachData);
                        File.WriteAllBytes(Path.Combine(outputDir, uniqueNum.ToString() + part.Filename), data);
                        uniqueNum++;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
            return count;
        }
        #endregion
    }
}
