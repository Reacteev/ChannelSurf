using ChannelSurfCli.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ChannelSurfCli.Utils
{
    public class Messages
    {
        public static void ScanMessagesByChannel(List<Models.Combined.ChannelsMapping> channelsMapping, string basePath,
            List<ViewModels.SimpleUser> slackUserList, String aadAccessToken, String selectedTeamId, bool copyFileAttachments, bool formatTimeToLocal,
            int nbMessagesPerFile, string fileAttachmentsPath, AttachmentIdFilePathMode attachmentIdMode, string jsonMessagesPath, string htmlMessagesPath)
        {
            foreach (var v in channelsMapping)
            {
                var channelAttachmentsToUpload = GetAndUploadMessages(v, basePath, slackUserList, aadAccessToken, selectedTeamId, copyFileAttachments, formatTimeToLocal, nbMessagesPerFile, fileAttachmentsPath, attachmentIdMode, jsonMessagesPath, htmlMessagesPath);
            }

            return;
        }


        static List<Models.Combined.AttachmentsMapping> GetAndUploadMessages(Models.Combined.ChannelsMapping channelsMapping, string basePath,
            List<ViewModels.SimpleUser> slackUserList, String aadAccessToken, String selectedTeamId, bool copyFileAttachments, bool formatTimeToLocal,
            int nbMessagesPerFile, string fileAttachmentsPath, AttachmentIdFilePathMode attachmentIdMode, string jsonMessagesPath, string htmlMessagesPath)
        {
            var messageList = new List<ViewModels.SimpleMessage>();
            messageList.Clear();

            var messageListJsonSource = new JArray();
            messageListJsonSource.Clear();

            List<Models.Combined.AttachmentsMapping> attachmentsToUpload = new List<Models.Combined.AttachmentsMapping>();
            attachmentsToUpload.Clear();

            Console.WriteLine("Migrating messages in channel " + channelsMapping.slackChannelName);
            foreach (var file in Directory.GetFiles(Path.Combine(basePath, channelsMapping.slackChannelName)))
            {
                Console.WriteLine("File " + file);
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                using (StreamReader sr = new StreamReader(fs))
                using (JsonTextReader reader = new JsonTextReader(sr))
                {
                    while (reader.Read())
                    {
                        if (reader.TokenType == JsonToken.StartObject)
                        {
                            JObject obj = JObject.Load(reader);

                            // SelectToken returns null not an empty string if nothing is found
                            // I'm too lazy right now for strongly typed classes

                            // deal with message basics: when, body, who

                            var messageTs = (string)obj.SelectToken("ts");
                            var messageText = (string)obj.SelectToken("text");
                            var messageId = channelsMapping.slackChannelId + "." + messageTs;
                            //messageText = RegexDetector.DetectSlackParens(messageText, slackUserList);
                            var messageSender = Utils.Messages.FindMessageSender(obj, slackUserList);

                            // create a list of attachments to upload
                            // deal with "attachments" that are files
                            // specifically, files hosted by Slack

                            var fileAttachments = new List<ViewModels.SimpleMessage.FileAttachment>();

                            var files = (JArray)obj.SelectToken("files");
                            if (files != null)
                            {
                                foreach (var attach in files)
                                {
                                    // SelectToken returns null not an empty string if nothing is found
                                    var fileUrl = (string)attach.SelectToken("url_private");
                                    var fileId = (string)attach.SelectToken("id");
                                    var fileMode = (string)attach.SelectToken("mode");
                                    var fileName = (string)attach.SelectToken("name");
                                    var local = (string)attach.SelectToken("url_private_download");
                                    if (local != null && local.StartsWith("_files"))
                                    {
                                        fileUrl = Path.Join(basePath, local);
                                    }

                                    if (fileMode != "external" && fileId != null && fileUrl != null)
                                    {
                                        Console.WriteLine("Message attachment found with ID " + fileId);
                                        attachmentsToUpload.Add(new Models.Combined.AttachmentsMapping
                                        {
                                            attachmentId = fileId,
                                            attachmentUrl = fileUrl,
                                            attachmentChannelId = channelsMapping.slackChannelId,
                                            attachmentFileName = fileName,
                                            msChannelName = channelsMapping.displayName
                                        });

                                        // map the attachment to fileAttachment which is used in the viewmodel
                                        fileAttachments.Add(new ViewModels.SimpleMessage.FileAttachment
                                        {
                                            id = fileId,
                                            originalName = (string)attach.SelectToken("name"),
                                            originalTitle = (string)attach.SelectToken("title"),
                                            originalUrl = (string)attach.SelectToken("permalink")
                                        });
                                    }
                                }
                            }
                            
                            // deal with "attachments" that aren't files

                            List<ViewModels.SimpleMessage.Attachments> attachmentsList = new List<ViewModels.SimpleMessage.Attachments>();
                            List<ViewModels.SimpleMessage.Attachments.Fields> fieldsList = new List<ViewModels.SimpleMessage.Attachments.Fields>();
                            var attachmentsObject = (JArray)obj.SelectToken("attachments");
                            if (attachmentsObject != null)
                            {

                                foreach (var attachmentItem in attachmentsObject)
                                {
                                    var attachmentText = (string)attachmentItem.SelectToken("text");
                                    var attachmentTextFallback = (string)attachmentItem.SelectToken("fallback");

                                    var attachmentItemToAdd = new ViewModels.SimpleMessage.Attachments();

                                    if (!String.IsNullOrEmpty(attachmentText))
                                    {
                                        attachmentItemToAdd.text = attachmentText;
                                    }
                                    else if (!String.IsNullOrEmpty(attachmentTextFallback))
                                    {
                                        attachmentItemToAdd.text = attachmentTextFallback;
                                    }

                                    var attachmentServiceName = (string)attachmentItem.SelectToken("service_name");
                                    if (!String.IsNullOrEmpty(attachmentServiceName))
                                    {
                                        attachmentItemToAdd.service_name = attachmentServiceName;
                                    }

                                    var attachmentFromUrl = (string)attachmentItem.SelectToken("from_url");
                                    if (!String.IsNullOrEmpty(attachmentFromUrl))
                                    {
                                        attachmentItemToAdd.url = attachmentFromUrl;
                                    }

                                    var attachmentColor = (string)attachmentItem.SelectToken("color");
                                    if (!String.IsNullOrEmpty(attachmentColor))
                                    {
                                        attachmentItemToAdd.color = attachmentColor;
                                    }

                                    var fieldsObject = (JArray)attachmentItem.SelectToken("fields");
                                    if (fieldsObject != null)
                                    {
                                        fieldsList.Clear();
                                        foreach (var fieldItem in fieldsObject)
                                        {
                                            fieldsList.Add(new ViewModels.SimpleMessage.Attachments.Fields()
                                            {
                                                title = (string)fieldItem.SelectToken("title"),
                                                value = (string)fieldItem.SelectToken("value"),
                                                shortWidth = (bool)fieldItem.SelectToken("short")
                                            });
                                        }
                                        attachmentItemToAdd.fields = fieldsList;
                                    }
                                    else
                                    {
                                        attachmentItemToAdd.fields = null;
                                    }
                                    attachmentsList.Add(attachmentItemToAdd);
                                }
                            }
                            else
                            {
                                attachmentsList = null;
                            }

                            // do some stuff with slack message threading at some point

                            messageList.Add(new ViewModels.SimpleMessage
                            {
                                id = messageId,
                                text = messageText,
                                ts = messageTs,
                                user = messageSender,
                                fileAttachments = fileAttachments,
                                attachments = attachmentsList,
                            });
                        }

                    }
                }
            }

            if (copyFileAttachments)
            {
                Utils.FileAttachments.ArchiveMessageFileAttachments(aadAccessToken, selectedTeamId, attachmentsToUpload, fileAttachmentsPath, attachmentIdMode).Wait();

                foreach (var messageItem in messageList)
                {
                    if (messageItem.fileAttachments.Count > 0)
                    {
                        foreach (var attach in messageItem.fileAttachments)
                        {
                            var messageItemWithFileAttachment = attachmentsToUpload.Find(w => String.Equals(attach.id, w.attachmentId, StringComparison.CurrentCultureIgnoreCase));
                            if (messageItemWithFileAttachment != null)
                            {
                                attach.spoId = messageItemWithFileAttachment.msSpoId;
                                attach.spoUrl = messageItemWithFileAttachment.msSpoUrl;
                            }
                        }
                    }
                }
            }
            Utils.Messages.CreateSlackMessageJsonArchiveFile(basePath, channelsMapping, messageList, aadAccessToken, selectedTeamId, nbMessagesPerFile, jsonMessagesPath);
            Utils.Messages.CreateSlackMessageHtmlArchiveFile(basePath, channelsMapping, messageList, aadAccessToken, selectedTeamId, slackUserList, formatTimeToLocal, nbMessagesPerFile, htmlMessagesPath);

            return attachmentsToUpload;
        }

        static void CreateSlackMessageJsonArchiveFile(String basePath, Models.Combined.ChannelsMapping channelsMapping, List<ViewModels.SimpleMessage> messageList,
            String aadAccessToken, string selectedTeamId, int nbMessagesPerFile, string jsonMessagesPath)
        {
            int messageIndexPosition = 0;

            if (nbMessagesPerFile < 0)
            {
                nbMessagesPerFile = messageList.Count;
            }

            for (int slackMessageFileIndex = 0; messageIndexPosition < messageList.Count; slackMessageFileIndex++)
            {
                var filenameToAdd = slackMessageFileIndex.ToString() + ".json";
                using (FileStream fs = new FileStream(Path.Combine(basePath, channelsMapping.slackChannelName, slackMessageFileIndex.ToString() + ".json"), FileMode.Create))
                {
                    using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                    {
                        int numOfMessagesToTake = 0;
                        if (messageIndexPosition + nbMessagesPerFile <= messageList.Count)
                        {
                            numOfMessagesToTake = nbMessagesPerFile;
                        }
                        else
                        {
                            numOfMessagesToTake = messageList.Count - messageIndexPosition;
                        }
                        var jsonObjectsToSave = JsonConvert.SerializeObject(messageList.Skip(messageIndexPosition).Take(numOfMessagesToTake), Formatting.Indented);
                        messageIndexPosition += numOfMessagesToTake;
                        w.WriteLine(jsonObjectsToSave);
                    }
                }
                var pathToItem = "/" + channelsMapping.displayName + jsonMessagesPath + filenameToAdd;
                Utils.FileAttachments.UploadFileToTeamsChannel(aadAccessToken, selectedTeamId, Path.Combine(basePath, channelsMapping.slackChannelName, filenameToAdd), pathToItem, false).Wait();
            }
            return;
        }

        static void CreateSlackMessageHtmlArchiveFile(String basePath, Models.Combined.ChannelsMapping channelsMapping, List<ViewModels.SimpleMessage> messageList,
            String aadAccessToken, string selectedTeamId, List<ViewModels.SimpleUser> slackUserList, bool formatTimeToLocal, int nbMessagesPerFile, string htmlMessagesPath)
        {
            int messageIndexPosition = 0;

            if (nbMessagesPerFile < 0)
            {
                nbMessagesPerFile = messageList.Count;
            }

            for (int slackMessageFileIndex = 0; messageIndexPosition < messageList.Count; slackMessageFileIndex++)
            {
                var filenameToAdd = slackMessageFileIndex.ToString() + ".html";
                using (FileStream fs = new FileStream(Path.Combine(basePath, channelsMapping.slackChannelName, slackMessageFileIndex.ToString() + ".html"), FileMode.Create))
                {
                    using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                    {
                        int numOfMessagesToTake = 0;
                        if (messageIndexPosition + nbMessagesPerFile <= messageList.Count)
                        {
                            numOfMessagesToTake = nbMessagesPerFile;
                        }
                        else
                        {
                            numOfMessagesToTake = messageList.Count - messageIndexPosition;
                        }
                        StringBuilder fileBody = new StringBuilder();
                        fileBody.Append(@"<html>
                            <head>
                            <style>
                            @import url('https://fonts.googleapis.com/css?family=Lato:400,900');
                            body {
	                            font-family: 'Lato', sans-serif;
                            }
                            </style>
                            </head>
                        ");
                        fileBody.Append("<body>");
                        fileBody.AppendLine("");
                        for (int i = 0; i < numOfMessagesToTake; i++)
                        {
                            var messageAsHtml = MessageToHtml(messageList[messageIndexPosition + i], channelsMapping, slackUserList, formatTimeToLocal);
                            fileBody.AppendLine(messageAsHtml);
                        }
                        fileBody.AppendLine("</body></html>");
                        messageIndexPosition += numOfMessagesToTake;
                        w.WriteLine(fileBody);
                    }
                }
                var pathToItem = "/" + channelsMapping.displayName + htmlMessagesPath + filenameToAdd;
                Utils.FileAttachments.UploadFileToTeamsChannel(aadAccessToken, selectedTeamId, Path.Combine(basePath, channelsMapping.slackChannelName, filenameToAdd), pathToItem, true).Wait();
            }

            return;
        }

        // this is ugly and should/will eventually be replaced by its own class

        private static string NiceDate(string epoch, bool formatTimeToLocal)
        {
            var seconds = float.Parse(epoch, System.Globalization.CultureInfo.InvariantCulture);
            var date = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(seconds);
            if (formatTimeToLocal)
            {
                date = date.ToLocalTime();
            }
            return date.ToString("dd/MM/yyyy HH:mm");
        }

        private static string NiceMention(string text, List<ViewModels.SimpleUser> slackUserList)
        {
            var nice = Regex.Replace(text, @"<@(\w+)>", m =>
            {
                var user = m.Groups[1].Value;
                if (user != "USLACKBOT")
                {
                    var simpleUser = slackUserList.FirstOrDefault(w => w.userId == user);
                    if (simpleUser != null)
                    {
                        return simpleUser.displayName();
                    }
                    else
                    {
                        return user;
                    }
                }
                else
                {
                    return "SlackBot";
                }
            });
            return nice;
        }

        private static string NiceEmoji(string text)
        {
            return text
                .Replace(":slightly_smiling_face:", "🙂")
                .Replace(":grinn", "😁")
                .Replace(":ghost", "👻");
        }

        private static string NiceText(string text, List<ViewModels.SimpleUser> slackUserList)
        {
            return NiceEmoji(NiceMention(text, slackUserList));
        }

        public static string MessageToHtml(ViewModels.SimpleMessage simpleMessage, Models.Combined.ChannelsMapping channelsMapping, List<ViewModels.SimpleUser> slackUserList, bool formatTimeToLocal)
        {
            string w = "";
            w += "<div>";
            w += ("<div id=\"" + simpleMessage.id + "\">");
            w += ("<span id=\"user_id\" style=\"font-weight:bold;\">" + simpleMessage.user + "</span>");
            w += ("&nbsp;");
            w += ("<span id=\"epoch_time\" style=\"font-weight:lighter;\">" + NiceDate(simpleMessage.ts, formatTimeToLocal) + "</span>");
            w += ("<br/>");
            w += ("<div id=\"message_text\" style=\"font-weight:normal;white-space:pre-wrap;\">" + NiceText(simpleMessage.text, slackUserList) + "</div>");

            if (simpleMessage.fileAttachments.Count > 0)
            {
                foreach (var attach in simpleMessage.fileAttachments)
                {
                    var isImg = attach.originalUrl.EndsWith(".jpg") || attach.originalUrl.EndsWith(".jpeg") || attach.originalUrl.EndsWith(".png") || attach.originalUrl.EndsWith(".gif");
                    w += "<div style=\"margin-left:1%;margin-top:1%;border-left-style:solid;border-left-color:LightGrey;\">";
                    w += "<div style=\"margin-left:1%;\">";
                    if (attach.spoId != null)
                    {
                        if (isImg)
                        {
                            w += "<img style='width:25%' src='" + attach.spoUrl + "' />";
                        }
                        else
                        {
                            w += "<span style=\"font-weight:lighter;\"> <a href=\"" + attach.spoUrl + "\"> File Attachment </a> </span>";
                        }
                    }
                    w += "<div>";
                    if (!isImg)
                    {
                        w += "<span style=\"font-weight:lighter;\"> ";
                        w += attach.originalTitle + "<br/>";
                        //w += attach.originalUrl + " <br/>";
                    }
                    w += "</span>";
                    w += "</div>";
                    w += "</div>";
                    w += "</div>";
                }
            }
            if (simpleMessage.attachments != null)
            {

                foreach (var attachment in simpleMessage.attachments)
                {
                    w += "<div style=\"margin-left:1%;margin-top:1%;border-left-style:solid;border-left-color:";
                    if (!String.IsNullOrEmpty(attachment.color))
                    {
                        w += "#" + attachment.color + ";";
                    }
                    else
                    {
                        w += "LightGrey;";
                    }
                    w += "\">";
                    w += "<div style=\"margin-left:1%;\">";
                    if (!String.IsNullOrEmpty(attachment.service_name))
                    {
                        w += "<span style=\"font-weight:bolder;\">" + attachment.service_name + "</span><br/>";
                    }
                    w += "<div style=\"font-weight:lighter;white-space:pre-wrap;\">" + attachment.text + "</div>";
                    w += "<a style=\"font-weight:lighter;\" href=\"" + attachment.url + "\">" + attachment.url + "</a><br/>";
                    if (attachment.fields != null)
                    {
                        if (attachment.fields.Count > 0)
                        {
                            w += "<table class=\"\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">";

                            foreach (var field in attachment.fields)
                            {
                                if (true)
                                {
                                    w += "<tr><td>";
                                    w += "<div>" + field.title + "</div>";
                                    w += "<div>" + field.value + "</div>";
                                    w += "</tr></td>";
                                }
                            }
                            w += "</table>";
                        }
                    }
                    w += "</div>";
                    w += "</div>";
                }
            }
            w += "</div>";
            w += "<p/>";
            w += "</div>";
            return w;
        }


        static string FindMessageSender(JObject obj, List<ViewModels.SimpleUser> slackUserList)
        {
            var user = (string)obj.SelectToken("user");
            if (!String.IsNullOrEmpty(user))
            {
                if (user != "USLACKBOT")
                {
                    var simpleUser = slackUserList.FirstOrDefault(w => w.userId == user);
                    if (simpleUser != null)
                    {
                        return simpleUser.displayName();
                    }

                }
                else
                {
                    return "SlackBot";
                }
            }
            else if (!(String.IsNullOrEmpty((string)obj.SelectToken("username"))))
            {
                return (string)obj.SelectToken("username");
            }
            else if (!(String.IsNullOrEmpty((string)obj.SelectToken("bot_id"))))
            {
                return (string)obj.SelectToken("bot_id");
            }

            return "";
        }
    }
}

