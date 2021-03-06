﻿using System;
using System.IO;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Globalization;
using System.Collections.Generic;
using ChannelSurfCli.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.EnvironmentVariables;
using System.Reflection;
using CommandLine;

namespace ChannelSurfCli
{
    public enum AttachmentIdFilePathMode
    {
        Directory,
        Prefix,
        Suffix
    }

    class Options
    {
        [Option("save-token", Default = false, HelpText = "Save auth token in local .token file to avoid going through auth flow.")]
        public bool SaveToken { get; set; }

        [Option("format-time-to-local", Default = false, HelpText = "Format messages date/time to local timezone.")]
        public bool FormatTimeToLocal { get; set; }

        [Option(Required = true, HelpText = "Input file to process, either channels.json or full Slack export.")]
        public string File { get; set; }

        [Option(HelpText = "Process only given channel.")]
        public string Only { get; set; }

        [Option("nb-messages-per-file", Default = "250", HelpText = "Number of messages to include in each html/json file. If all or negative value provided, all messages will be exported in a single file.")]
        public string NbMessagesPerFile { get; set; }

        [Option("sharepoint-path", Default = "channelsurf", HelpText = "Base path to stored migrated files (both messages and attachments).")]
        public string SharePointPath { get; set; }

        [Option("attachment-id-mode", Default = AttachmentIdFilePathMode.Directory, HelpText = "Way to put attachment Id in file path. Could be \n  - 'Directory' (default) to create a sub-directory for each attachment Id\n  - Prefix to put Id as prefix in filename\n  - Suffix to put Id as suffix in filename (before extension)")]
        public AttachmentIdFilePathMode AttachmentIdMode { get; set; }

        [Option("file-attachments-path", Default = "fileattachments", HelpText = "Sub-path inside SharePointPath to store file attachments to.")]
        public string FileAttachmentsPath { get; set; }

        [Option("messages-path", Default = "messages", HelpText = "Sub-path inside SharePointPath to store messages to.")]
        public string MessagesPath { get; set; }

        [Option("html-messages-path", Default = "html", HelpText = "Sub-path inside SharePointPath/MessagesPath to store html messages to.")]
        public string HtmlMessagesPath { get; set; }

        [Option("json-messages-path", Default = "json", HelpText = "Sub-path inside SharePointPath/MessagesPath to store json messages to.")]
        public string JsonMessagesPath { get; set; }
    }

    class Program
    {
        // all of your per-tenant and per-environment settings are (now) in appsettings.json
        public static IConfigurationRoot Configuration { get; set; }

        // Don't change this constant
        // It is a constant that corresponds to fixed values in AAD that corresponds to Microsoft Graph

        // Required Permissions - Microsoft Graph -> API
        // Read all users' full profiles
        // Read and write all groups

        const string aadResourceAppId = "00000003-0000-0000-c000-000000000000";

        static AuthenticationContext authenticationContext = null;
        static AuthenticationResult authenticationResult = null;

        static void Main(string[] args)
        {
            string slackArchiveBasePath = "";
            string slackArchiveTempPath = "";
            string channelsPath = "";
            bool channelsOnly = false;
            bool copyFileAttachments = false;
            int nbMessagesPerFile = 250;

            Console.CancelKeyPress += delegate (object sender, ConsoleCancelEventArgs e)
            {
                if (!channelsOnly)
                {
                    try
                    {
                        Utils.Files.CleanUpTempDirectoriesAndFiles(slackArchiveTempPath);
                    }
                    catch
                    {
                        // to-do: something 
                    }
                }
            };

            var result = Parser.Default.ParseArguments<Options>(args)
                .WithParsed(options =>
                {
                    // retreive settings from appsettings.json instead of hard coding them here

                    var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                        .AddEnvironmentVariables();
                    Configuration = builder.Build();

                    // Parse command line options

                    var sharepointPath = options.SharePointPath;
                    if (! sharepointPath.StartsWith("/"))
                    {
                        sharepointPath = "/" + sharepointPath;
                    }
                    if (sharepointPath.EndsWith("/"))
                    {
                        sharepointPath = sharepointPath.Substring(0, sharepointPath.Length - 1);
                    }

                    var fileAttachmentsPath = options.FileAttachmentsPath;
                    if (! fileAttachmentsPath.StartsWith("/"))
                    {
                        fileAttachmentsPath = "/" + fileAttachmentsPath;
                    }
                    fileAttachmentsPath = sharepointPath + fileAttachmentsPath;
                    if (! fileAttachmentsPath.EndsWith("/"))
                    {
                        fileAttachmentsPath = fileAttachmentsPath + "/";
                    }

                    var messagesPath = options.MessagesPath;
                    if (! messagesPath.StartsWith("/"))
                    {
                        messagesPath = "/" + messagesPath;
                    }
                    messagesPath = sharepointPath + messagesPath;
                    if (! messagesPath.EndsWith("/"))
                    {
                        messagesPath = messagesPath + "/";
                    }

                    var htmlMessagesPath = options.HtmlMessagesPath;
                    if (htmlMessagesPath.StartsWith("/"))
                    {
                        htmlMessagesPath = htmlMessagesPath.Substring(0, htmlMessagesPath.Length - 1);
                    }
                    htmlMessagesPath = messagesPath + htmlMessagesPath;
                    if (! htmlMessagesPath.EndsWith("/"))
                    {
                        htmlMessagesPath = htmlMessagesPath + "/";
                    }

                    var jsonMessagesPath = options.JsonMessagesPath;
                    if (jsonMessagesPath.StartsWith("/"))
                    {
                        jsonMessagesPath = jsonMessagesPath.Substring(0, jsonMessagesPath.Length - 1);
                    }
                    jsonMessagesPath = messagesPath + jsonMessagesPath;
                    if (! jsonMessagesPath.EndsWith("/"))
                    {
                        jsonMessagesPath = jsonMessagesPath + "/";
                    }

                    // Start processing

                    Console.WriteLine("");
                    Console.WriteLine("****************************************************************************************************");
                    Console.WriteLine("Welcome to Channel Surf!");
                    Console.WriteLine("This tool makes it easy to bulk create channels in an existing Microsoft Team.");
                    Console.WriteLine("All we need a Slack Team export ZIP file whose channels you wish to re-create.");
                    Console.WriteLine("Or, you can define new channels in a file called channels.json.");
                    Console.WriteLine("****************************************************************************************************");
                    Console.WriteLine("");

                    while (Configuration["AzureAd:TenantId"] == "" || Configuration["AzureAd:ClientId"] == "")
                    {
                        Console.WriteLine("");
                        Console.WriteLine("****************************************************************************************************");
                        Console.WriteLine("You need to provide your Azure Active Directory Tenant Name and the Application ID you created for");
                        Console.WriteLine("use with application to continue.  You can do this by altering Program.cs and re-compiling this app.");
                        Console.WriteLine("Or, you can provide it right now.");
                        Console.Write("Azure Active Directory Tenant Name (i.e your-domain.onmicrosoft.com): ");
                        Configuration["AzureAd:TenantId"] = Console.ReadLine();
                        Console.Write("Azure Active Directory Application ID: ");
                        Configuration["AzureAd:ClientId"] = Console.ReadLine();
                        Console.WriteLine("****************************************************************************************************");
                    }

                    Console.WriteLine("**************************************************");
                    Console.WriteLine("Tenant is " + (Configuration["AzureAd:TenantId"]));
                    Console.WriteLine("Application ID is " + (Configuration["AzureAd:ClientId"]));
                    Console.WriteLine("Redirect URI is " + (Configuration["AzureAd:AadRedirectUri"]));
                    Console.WriteLine("**************************************************");

                    Console.WriteLine("");
                    Console.WriteLine("****************************************************************************************************");
                    Console.WriteLine("Your tenant admin consent URL is https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token" +
                        "&client_id=" + Configuration["AzureAd:ClientId"] + "&redirect_uri=" + Configuration["AzureAd:AadRedirectUri"] + "&prompt=admin_consent" + "&nonce=" + Guid.NewGuid().ToString());
                    Console.WriteLine("****************************************************************************************************");
                    Console.WriteLine("");


                    Console.WriteLine("");
                    Console.WriteLine("****************************************************************************************************");
                    Console.WriteLine("Let's get started! Sign in to Microsoft with your Teams credentials:");

                    var aadAccessToken = "";
                    var tokenExists = File.Exists(".token");
                    if (options.SaveToken && tokenExists)
                    {
                        aadAccessToken = File.ReadAllText(".token");
                    }
                    else
                    {
                        authenticationResult = UserLogin();
                        aadAccessToken = authenticationResult.AccessToken;
                        if (String.IsNullOrEmpty(aadAccessToken))
                        {
                            Console.WriteLine("Something went wrong.  Please try again!");
                            Environment.Exit(1);
                        }
                        else
                        {
                            Console.WriteLine("You've successfully signed in.  Welcome " + authenticationResult.UserInfo.DisplayableId);
                        }
                        if (options.SaveToken)
                        {
                            File.WriteAllText(".token", aadAccessToken);
                        }
                    }

                    var selectedTeamId = Utils.Channels.SelectJoinedTeam(aadAccessToken);
                    if (selectedTeamId == "")
                    {
                        if (options.SaveToken && tokenExists)
                        {
                            File.Delete(".token");
                        }
                        Environment.Exit(0);
                    }

                    if (options.File.EndsWith("channels.json", StringComparison.CurrentCulture))
                    {
                        channelsPath = options.File;
                        channelsOnly = true;
                    }
                    else
                    {
                        slackArchiveTempPath = Path.GetTempFileName();
                        slackArchiveBasePath = Utils.Files.DecompressSlackArchiveFile(options.File, slackArchiveTempPath);
                        channelsPath = Path.Combine(slackArchiveBasePath, "channels.json");
                    }

                    Console.WriteLine("Scanning channels.json");
                    var slackChannelsToMigrate = Utils.Channels.ScanSlackChannelsJson(channelsPath);
                    if (!string.IsNullOrEmpty(options.Only))
                    {
                        slackChannelsToMigrate = slackChannelsToMigrate.FindAll(c => c.channelName == options.Only);
                    }

                    Console.WriteLine("Creating channels in MS Teams");
                    var msTeamsChannelsWithSlackProps = Utils.Channels.CreateChannelsInMsTeams(aadAccessToken, selectedTeamId, slackChannelsToMigrate, slackArchiveTempPath, sharepointPath);
                    Console.WriteLine("Creating channels in MS Teams - done");

                    if (channelsOnly)
                    {
                        Environment.Exit(0);
                    }

                    if (!string.IsNullOrEmpty(options.NbMessagesPerFile))
                    {
                        if (options.NbMessagesPerFile.Equals("all", StringComparison.CurrentCultureIgnoreCase))
                        {
                            nbMessagesPerFile = -1;
                        }
                        else
                        {
                            nbMessagesPerFile = int.Parse(options.NbMessagesPerFile);
                        }
                    }

                    Console.Write("Create web pages that show the message history for each re-created Slack channel? (y|n): ");
                    var copyMessagesResponse = Console.ReadLine();
                    if (copyMessagesResponse.StartsWith("y", StringComparison.CurrentCultureIgnoreCase))
                    {
                        Console.Write("Copy files attached to Slack messages to Microsoft Teams? (y|n): ");
                        var copyFileAttachmentsResponse = Console.ReadLine();
                        if (copyFileAttachmentsResponse.StartsWith("y", StringComparison.CurrentCultureIgnoreCase))
                        {
                            copyFileAttachments = true;
                        }

                        Console.WriteLine("Scanning users in Ms Teams organization");
                        var msTeamsUserList = Utils.O365.getUsers(aadAccessToken);
                        Console.WriteLine("Scanning users in Ms Teams organization - done");

                        Console.WriteLine("Scanning users in Slack archive");
                        var slackUserList = Utils.Users.ScanUsers(Path.Combine(slackArchiveBasePath, "users.json"), msTeamsUserList);
                        Console.WriteLine("Scanning users in Slack archive - done");

                        Console.WriteLine("Scanning messages in Slack channels");
                        Utils.Messages.ScanMessagesByChannel(msTeamsChannelsWithSlackProps, slackArchiveTempPath, slackUserList, aadAccessToken, selectedTeamId, copyFileAttachments, options.FormatTimeToLocal, nbMessagesPerFile, fileAttachmentsPath, options.AttachmentIdMode, jsonMessagesPath, htmlMessagesPath);
                        Console.WriteLine("Scanning messages in Slack channels - done");
                    }

                    Console.WriteLine("Tasks complete.  Press any key to exit");
                    Console.ReadKey();

                    Utils.Files.CleanUpTempDirectoriesAndFiles(slackArchiveTempPath);
                })
                .WithNotParsed(_ =>
                {
                    Console.WriteLine("Usage: channelsurf --file export.zip [--only channelname]");
                });
        }

        static AuthenticationResult UserLogin()
        {
            authenticationContext = new AuthenticationContext
                    (String.Format(CultureInfo.InvariantCulture, Configuration["AzureAd:AadInstance"], Configuration["AzureAd:TenantId"]));
            authenticationContext.TokenCache.Clear();
            DeviceCodeResult deviceCodeResult = authenticationContext.AcquireDeviceCodeAsync(aadResourceAppId, (Configuration["AzureAd:ClientId"])).Result;
            Console.WriteLine(deviceCodeResult.Message);
            return authenticationContext.AcquireTokenByDeviceCodeAsync(deviceCodeResult).Result;
        }

    }
}
