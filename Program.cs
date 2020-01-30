using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;
using Discord;
using Discord.WebSocket;
using Quartz;
using Quartz.Impl;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace Bible_Bot
{
    public class Program
    {
    //https://discord.foxbot.me/docs/api/index.html
    private DiscordSocketClient _client;
    private ulong _channel = ulong.Parse(ConfigurationManager.AppSettings["Discord_ChannelKey"]);
    private ulong _guild = ulong.Parse(ConfigurationManager.AppSettings["Discord_GuildKey"]);
    static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
    static string ApplicationName = "Bible Bot";

	public static void Main(string[] args)
		=> new Program().MainAsync().GetAwaiter().GetResult();

	public async Task MainAsync()
	{
        SheetsService service = EstablishSpreadsheet();

        _client = new DiscordSocketClient();
        _client.Log += Log;

        // Begin Discord connection
        await _client.LoginAsync(TokenType.Bot, ConfigurationManager.AppSettings["Discord_OAuth2Token"]);
        await _client.StartAsync();

        // Establish Quartz scheduler
        NameValueCollection props = new NameValueCollection
            {
                { "quartz.serializer.type", "binary" }
            };
        StdSchedulerFactory factory = new StdSchedulerFactory(props);
        IScheduler scheduler = await factory.GetScheduler();

        scheduler.Context.Put("service", service);
        scheduler.Context.Put("dcordClient", _client);

        // and start it off
        await scheduler.Start();
        
        IJobDetail biblePassageJob = JobBuilder.Create<MessageBiblePassageJob>()
        .WithIdentity("daily", "group1")
        .UsingJobData("channelKey", _channel.ToString())
        .UsingJobData("guildKey", _guild.ToString())
        .Build();
        
        // https://www.quartz-scheduler.net/documentation/quartz-3.x/tutorial/crontrigger.html
        ITrigger trigger = TriggerBuilder.Create()
        .WithIdentity("daily", "group1")
        .WithCronSchedule("0 0 " + ConfigurationManager.AppSettings["Quartz_ScheduledHour"] + " * * ? *")
        .ForJob(biblePassageJob)
        .Build();

        await scheduler.ScheduleJob(biblePassageJob, trigger);

        await Task.Delay(-1);
	}

    private SheetsService EstablishSpreadsheet()
    {
        UserCredential credential;

        using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        {
            // The file token.json stores the user's access and refresh tokens, and is created
            // automatically when the authorization flow completes for the first time.
            string credPath = "token.json";
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                Scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
            Console.WriteLine("Credential file saved to: " + credPath);
        }

        // Create Google Sheets API service.
        SheetsService service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });

        return service;
    }

    private Task Log(LogMessage msg)
    {
        Console.WriteLine(msg.ToString());
	    return Task.CompletedTask;
    }

    }
    
    public class MessageBiblePassageJob : IJob
    {
        public async Task Execute(IJobExecutionContext context)
	    {
            JobDataMap dataMap = context.JobDetail.JobDataMap;
            DiscordSocketClient dcordClient = (DiscordSocketClient)context.Scheduler.Context.Get("dcordClient");
            SheetsService service = (SheetsService)context.Scheduler.Context.Get("service");
            ulong clientKey = ulong.Parse(dataMap.GetString("channelKey"));
            ulong guildKey = ulong.Parse(dataMap.GetString("guildKey"));
            SocketTextChannel channel = dcordClient.GetGuild(guildKey).GetChannel(clientKey) as SocketTextChannel;

            await Console.Out.WriteLineAsync("Job executing");

            // Define request parameters.
            String spreadsheetId = ConfigurationManager.AppSettings["Google_SpreadsheetID"];
            String range = String.Format("'{0}/{1}'!A{2}:C{2}", DateTime.Today.Month, DateTime.Today.Year, DateTime.Today.Day + 1);
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Print
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            String msgString = "";
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    if (row.Count == 3)
                    {
                        msgString = String.Format("Hello! The reading for today ({0}) is from {1} {2}.\nhttps://www.biblegateway.com/passage/?search={1}+{2}&version=NKJV\nGod bless!", DateTime.Today.DayOfWeek.ToString(), row[1], row[2]);
                    }
                }
            }
            
            if (msgString != "")
            {
                await channel.SendMessageAsync(msgString);
            }
	    }
    }
}
