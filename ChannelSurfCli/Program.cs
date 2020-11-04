using System;
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
    class Options
    {
        [Option(Default = false, HelpText = "Save auth token in local .token file to avoid going through auth flow.")]
        public bool SaveToken { get; set; }

        [Option(Default = false, HelpText = "Format messages date/time to local timezone.")]
        public bool FormatTimeToLocal { get; set; }

        [Option(Required = true, HelpText = "Input file to process, either channels.json or full Slack export.")]
        public string File { get; set; }

        [Option(HelpText = "Process only given channel.")]
        public string Only { get; set; }

        [Option(Default = "250", HelpText = "Number of messages to include in each html/json file. If all or negative value provided, all messages will be exported in a single file.")]
        public string NbMessagesPerFile { get; set; }
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
                    if (options.SaveToken && File.Exists(".token"))
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
                    var msTeamsChannelsWithSlackProps = Utils.Channels.CreateChannelsInMsTeams(aadAccessToken, selectedTeamId, slackChannelsToMigrate, slackArchiveTempPath);
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

                        Console.WriteLine("Scanning users in Slack archive");
                        var slackUserList = Utils.Users.ScanUsers(Path.Combine(slackArchiveBasePath, "users.json"));
                        Console.WriteLine("Scanning users in Slack archive - done");

                        Console.WriteLine("Scanning messages in Slack channels");
                        Utils.Messages.ScanMessagesByChannel(msTeamsChannelsWithSlackProps, slackArchiveTempPath, slackUserList, aadAccessToken, selectedTeamId, copyFileAttachments, options.FormatTimeToLocal, nbMessagesPerFile);
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
