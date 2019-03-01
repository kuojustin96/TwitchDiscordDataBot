using System;
using System.Linq;
using System.Timers;
using System.Diagnostics;
using System.Collections.Generic;
using TwitchLib.Client;
using TwitchLib.Client.Models;
using TwitchLib.Client.Events;
using TwitchLib.Api;
using TwitchLib.Api.V5.Models.Users;
using Microsoft.Office.Interop.Excel;
using TwitchLib.Api.Services;
using TwitchLib.Api.Services.Events.LiveStreamMonitor;
using TwitchLib.Api.Services.Events.FollowerService;
using TwitchChatDataBot.Discord_Modules;

namespace TwitchChatDataBot
{
    internal class TwitchChatBot
    {
        readonly ConnectionCredentials credentials = new ConnectionCredentials(TwitchInfo.BotUsername, TwitchInfo.BotToken);
        TwitchClient client;
        TwitchAPI api;
        LiveStreamMonitorService monitor;
        FollowerService fservice;

        public Application excel;
        public Workbook wb;
        Worksheet ws;

        Stopwatch streamTrackingStopwatch = new Stopwatch();
        public bool autoDisplayData = true;

        //Total Stream Variables - Organize these
        //Need to reset these on stream start
        int maxViewers = 0;
        int oldViewCount = 0;
        int totalViewers = 0;
        List<int> viewsList = new List<int>();
        int lastViewCount = 0;
        int followersGained = 0;
        bool followersFirstTime = true;
        int subscribersGained = 0;
        int totalMessagesSent = 0;
        Dictionary<string, int> emoteDict = new Dictionary<string, int>();
        int totalNumEmotes = 0;


        //Interval Variables
        public int lastIntervalRow = 2;
        int int_intervalNum = 0;
        List<int> int_viewList = new List<int>();
        int int_numMessagesSent = 0;
        int int_numEmotesSent = 0;
        Dictionary<string, int> int_emoteDict = new Dictionary<string, int>();
        int int_followersGained = 0;
        int int_subscribersGained = 0;

        public TwitchChatBot()
        {
            //ORGANIZE THIS SHIT DNIAOSDNOAIDNIAO FUNCTIONS
        }

        #region TwitchLib Events
        internal void Connect()
        {
            Console.WriteLine("Connecting...");

            client = new TwitchClient();
            client.Initialize(credentials, TwitchInfo.ChannelName);

            //Client Events
            client.OnLog += Client_OnLog;
            client.OnConnectionError += Client_OnConnectionError;
            client.OnMessageReceived += Client_OnMessageReceived;
            client.OnNewSubscriber += Client_OnNewSubscriber;

            client.Connect();

            api = new TwitchAPI();
            api.Settings.ClientId = TwitchInfo.ClientID;
            api.Settings.AccessToken = TwitchInfo.BotToken;


            //API Events
            monitor = new LiveStreamMonitorService(api, int.Parse(TwitchInfo.UpdateInterval));
            List<string> channels = new List<string>();
            channels.Add(TwitchInfo.ChannelName);
            monitor.SetChannelsByName(channels);
            monitor.OnStreamOffline += Monitor_OnStreamOffline;
            monitor.OnStreamOnline += Monitor_OnStreamOnline;
            monitor.Start();

            fservice = new FollowerService(api, int.Parse(TwitchInfo.UpdateInterval));
            fservice.SetChannelsByName(channels);
            fservice.OnNewFollowersDetected += Fservice_OnNewFollowersDetected;
            fservice.Start();

            //Excel Stuff
            excel = new Application();
            excel.Visible = true;
            wb = excel.Workbooks.Open(TwitchInfo.ExcelPath);

            foreach (Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == TwitchInfo.ChannelName)
                {
                    Console.WriteLine("Found exisiting channel...");
                    ws = sheet;
                    break;
                }
            }

            if (ws == null)
            {
                //Create/copy a new worksheet from base worksheet
                Console.WriteLine("New channel detected, creating a new sheet...");
                ws = (Worksheet)excel.Worksheets.Add();
                ws.Name = TwitchInfo.ChannelName;
            }
        }

        internal void Disconnect()
        {

            CloseExcel();
            client.Disconnect();
            monitor.Stop();
            fservice.Stop();
            Console.WriteLine("Disconnecting...");
        }


        //Monitor Events
        private void Monitor_OnStreamOnline(object sender, OnStreamOnlineArgs e)
        {
            SetUpNewDataCollection();

            followersFirstTime = true;

            streamTrackingStopwatch.Reset();
            streamTrackingStopwatch.Start();

            ViewCountTimer(true);
            StreamDataIntervalTimer(true);

            totalViewers = e.Stream.ViewerCount;
            oldViewCount = totalViewers;

            Console.WriteLine("Stream is online with " + e.Stream.ViewerCount + " viewers");
        }

        private void Monitor_OnStreamOffline(object sender, OnStreamOfflineArgs e)
        {
            //SET TOTAL VALUES HERE AS WELL
            //END INTERVAL TIMER, ADD IT TO END

            SetIntervalData();
            SetTotalValues();

            if (streamTrackingStopwatch.IsRunning)
            {
                streamTrackingStopwatch.Stop();
                string temp = streamTrackingStopwatch.Elapsed.ToString();
                ws.Cells[lastIntervalRow + 15, 13] = temp;
            }

            streamTrackingStopwatch.Stop();
            string streamTrackingDuration = streamTrackingStopwatch.Elapsed.ToString();
            ws.Cells[17, 13] = streamTrackingDuration;

            ViewCountTimer(false);
            StreamDataIntervalTimer(false);
            
            Console.WriteLine("Stream is offline");
        }


        //Client Events
        private void Client_OnMessageReceived(object sender, OnMessageReceivedArgs e)
        {
            totalMessagesSent++;
            int_numMessagesSent++;

            if (e.ChatMessage.EmoteSet.Emotes.Count > 0)
            {
                foreach (EmoteSet.Emote emote in e.ChatMessage.EmoteSet.Emotes)
                {
                    if (emoteDict.ContainsKey(emote.Name))
                        emoteDict[emote.Name]++;
                    else
                        emoteDict.Add(emote.Name, 1);

                    if (int_emoteDict.ContainsKey(emote.Name))
                        int_emoteDict[emote.Name]++;
                    else
                        int_emoteDict.Add(emote.Name, 1);
                }

                totalNumEmotes += e.ChatMessage.EmoteSet.Emotes.Count;
                int_numEmotesSent += e.ChatMessage.EmoteSet.Emotes.Count;
            }
        }

        private void Client_OnNewSubscriber(object sender, OnNewSubscriberArgs e)
        {
            subscribersGained++;
            int_subscribersGained++;
            Console.WriteLine("New subsciber: " + e.Subscriber.DisplayName);
        }

        private void Client_OnConnectionError(object sender, OnConnectionErrorArgs e)
        {
            Console.WriteLine($"Error! {e.Error}");
        }

        private void Client_OnLog(object sender, OnLogArgs e)
        {
            //Console.WriteLine(e.Data);
        }


        //Follower Service
        private void Fservice_OnNewFollowersDetected(object sender, OnNewFollowersDetectedArgs e)
        {
            if (followersFirstTime)
            {
                followersFirstTime = false;
                return;
            }


            Console.WriteLine("New Followers Detected: " + e.NewFollowers.Count);
            followersGained += e.NewFollowers.Count;
            int_followersGained += e.NewFollowers.Count;
        }

        #endregion


        void SetUpNewDataCollection()
        {
            Console.WriteLine("Setting new data collection!");
            Worksheet template = null;
            foreach (Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == "Template Worksheet")
                {
                    template = sheet;
                    break;
                }
            }

            if (template == null)
                return;

            ws.Cells[3, 16].Value2 += 1;

            int rowNum = -1;
            foreach (Range r in ws.Rows)
            {
                if (ws.Cells[r.Row, 1].Value2 != null && ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                {
                    lastIntervalRow = r.Row;
                    //Console.WriteLine("LastIntervalRow: " + lastIntervalRow);
                }

                if (ws.Cells[r.Row, 1].Value2 == null)
                {
                    rowNum = r.Row;
                    //Console.WriteLine("RowNum: " + rowNum);
                    break;
                }
            }

            if (rowNum > 2)
            {
                if (rowNum < lastIntervalRow + 16)
                {
                    int difference = (lastIntervalRow + 16) - rowNum;
                    for (int x = 0; x < difference + 3; x++)
                        ws.Cells[rowNum + x, 1] = "_";
                } else
                {
                    for (int x = 0; x < 3; x++)
                        ws.Cells[rowNum + x, 1] = "_";
                }
            }

            Range from = template.UsedRange;
            if (rowNum == 1)
            {
                from.Copy(ws.Cells[rowNum, 1]);
            } else
            {
                if (rowNum < lastIntervalRow + 16)
                {
                    from.Copy(ws.Cells[lastIntervalRow + 19, 1]);
                    lastIntervalRow = lastIntervalRow + 20;
                }
                else
                {
                    from.Copy(ws.Cells[rowNum + 3, 1]);
                    lastIntervalRow = rowNum + 4;
                }
            }

            ws.Cells[lastIntervalRow - 1, 1] = string.Format("Stream {0} tracked on {1}", ws.Cells[3, 16].Value2, DateTime.Now.ToString());
        }

        private void ViewCountTimer(bool enable)
        {
            Timer timer = new Timer(int.Parse(TwitchInfo.UpdateInterval) * 1000);
            if (enable)
            {
                timer.Elapsed += OnViewerTimerEvent;
                timer.AutoReset = true;
                timer.Enabled = true;
            } else
            {
                timer.Elapsed -= OnViewerTimerEvent;
                timer.Stop();
                timer.Enabled = false;
            }
        }

        private void OnViewerTimerEvent(object sender, ElapsedEventArgs e)
        {
            if (monitor.LiveStreams.ContainsKey(TwitchInfo.ChannelName))
            {
                int viewerCount = monitor.LiveStreams[TwitchInfo.ChannelName].ViewerCount;

                if (viewerCount != lastViewCount)
                    viewsList.Add(viewerCount);

                if (int_viewList.Count == 0)
                    int_viewList.Add(viewerCount);

                else if (viewerCount != lastViewCount)
                    int_viewList.Add(viewerCount);
                Console.WriteLine("Number of viewers: " + viewerCount);

                //Most Concurrent
                if (viewerCount > maxViewers)
                    maxViewers = viewerCount;

                //Total Viewers
                if(viewerCount > oldViewCount)
                {
                    int difference = viewerCount - oldViewCount;
                    Console.WriteLine("Difference: " + difference);
                    totalViewers += difference;
                }

                //ehhhhhh... good enough
                oldViewCount = viewerCount;
            }
        }

        private void StreamDataIntervalTimer(bool enable)
        {
            Timer timer = new Timer(int.Parse(TwitchInfo.StreamDataInterval));
            if (enable)
            {
                timer.Elapsed += OnStreamDataTimerEvent;
                timer.AutoReset = true;
                timer.Enabled = true;
            }
            else
            {
                timer.Elapsed -= OnStreamDataTimerEvent;
                timer.Stop();
                timer.Enabled = false;
                timer.Dispose();
            }
        }

        private void OnStreamDataTimerEvent(object sender, ElapsedEventArgs e)
        {
            Console.WriteLine("Stream Data Timer Event Fired");

            SetIntervalData();
        }

        private void SetIntervalData()
        {
            //Find empty row
            int rowNum = -1;
            foreach (Range r in ws.Rows)//Errors here sometimes, could just a debugging error on disconnect
            {
                if (ws.Cells[r.Row, 1].Value2 == null)
                {
                    rowNum = r.Row;
                    break;
                }
            }

            int_intervalNum++;
            ws.Cells[rowNum, 1] = int_intervalNum;
            int avg = 0;
            if (int_viewList.Count > 0)
            {
                for (int x = 0; x < int_viewList.Count; x++)
                    avg += int_viewList[x];
                avg = avg / int_viewList.Count;
            }

            ws.Cells[rowNum, 2] = avg;
            ws.Cells[rowNum, 3] = int_numMessagesSent;
            ws.Cells[rowNum, 4] = int_numEmotesSent;
            string topEmotes = GetTopEmotes(int_emoteDict);
            ws.Cells[rowNum, 5] = GetTopEmotes(int_emoteDict);
            ws.Cells[rowNum, 6] = int_followersGained;
            ws.Cells[rowNum, 7] = int_subscribersGained;

            //Updates the total averages
            SetTotalValues();

            string streamTrackingDuration = streamTrackingStopwatch.Elapsed.ToString();
            ws.Cells[lastIntervalRow + 15, 13].NumberFormat = "HH:mm:ss";
            ws.Cells[lastIntervalRow + 15, 13] = streamTrackingDuration;

            ResetIntervalVariables();
        }

        private void ResetIntervalVariables()
        {
            int_viewList.Clear();
            int_numMessagesSent = 0;
            int_numEmotesSent = 0;
            int_emoteDict.Clear();
            int_followersGained = 0;
            int_subscribersGained = 0;
        }

        private TimeSpan? GetUptime()
        {
            string userID = GetUserID(TwitchInfo.ChannelName);
            if (userID == null) {
                Console.WriteLine("ERROR: Count not get Channel ID");
                return null;
            }

            return api.V5.Streams.GetUptimeAsync(userID).Result;
        }

        private string GetUserID(string username)
        {
            User[] userList = api.V5.Users.GetUserByNameAsync(username).Result.Matches;
            if (userList == null || userList.Length == 0)
                return null;
            else
                return userList[0].Id;
        }

        private string GetTopEmotes(Dictionary<string, int> dict)
        {
            List<KeyValuePair<string, int>> temp = dict.ToList().OrderBy(o => o.Value).ToList();
            string topEmotes = "";
            int topEmoteLength = int.Parse(TwitchInfo.TopEmotesLength);
            if (temp.Count > topEmoteLength)
            {
                for (int x = 0; x < topEmoteLength; x++)
                    topEmotes += string.Format("[{0}, {1}] ", temp[temp.Count - 1 - x].Key, temp[temp.Count - 1 - x].Value);
            }
            else
            {
                for (int x = temp.Count - 1; x > 0; x--)
                    topEmotes += string.Format("[{0}, {1}] ", temp[x].Key, temp[x].Value);
            }

            return topEmotes;
        }

        public void SetTotalValues()
        {
            //Search for total values section
            ws.Cells[lastIntervalRow + 2, 13] = totalViewers;
            ws.Cells[lastIntervalRow + 3, 13] = maxViewers;

            int avgViews = 0;
            if (viewsList.Count > 0)
            {
                for (int x = 0; x < viewsList.Count; x++)
                    avgViews += viewsList[x];
                avgViews = avgViews / viewsList.Count;
            }

            ws.Cells[lastIntervalRow + 4, 13] = avgViews;
            ws.Cells[lastIntervalRow + 5, 13] = followersGained;
            ws.Cells[lastIntervalRow + 6, 13] = subscribersGained;
            ws.Cells[lastIntervalRow + 7, 13] = totalNumEmotes;
            ws.Cells[lastIntervalRow + 8, 13] = GetTopEmotes(emoteDict);
            ws.Cells[lastIntervalRow + 9, 13] = totalMessagesSent;

            int avgMsg = 0;
            int avgEmotesUsed = 0;
            int avgFollows = 0;
            int avgSubs = 0;
            if (int_intervalNum > 0)
            {
                for (int x = 0; x < int_intervalNum; x++)
                {
                    avgMsg += ws.Cells[lastIntervalRow + 1 + x, 3].Value2;
                    avgEmotesUsed += ws.Cells[lastIntervalRow + 1 + x, 4].Value2;
                    avgFollows += ws.Cells[lastIntervalRow + 1 + x, 6].Value2;
                    avgSubs += ws.Cells[lastIntervalRow + 1 + x, 7].Value2;
                }

                ws.Cells[lastIntervalRow + 10, 13] = avgMsg / int_intervalNum;
                ws.Cells[lastIntervalRow + 11, 13] = avgEmotesUsed / int_intervalNum;
                ws.Cells[lastIntervalRow + 12, 13] = avgFollows / int_intervalNum;
                ws.Cells[lastIntervalRow + 13, 13] = avgSubs / int_intervalNum;
            } else
            {
                ws.Cells[lastIntervalRow + 10, 13] = avgMsg;
                ws.Cells[lastIntervalRow + 11, 13] = avgEmotesUsed;
                ws.Cells[lastIntervalRow + 12, 13] = avgFollows;
                ws.Cells[lastIntervalRow + 13, 13] = avgSubs;
            }

            ws.Cells[lastIntervalRow + 14, 13].NumberFormat = "HH:mm:ss";
            ws.Cells[lastIntervalRow + 14, 13] = GetUptime().ToString();
        }

        internal void CloseExcel()
        {
            //Not closing excel properly good job
            //Need to add this to on stream offline
            SetIntervalData(); //TEMPORARY, maybe keep here as well
            SetTotalValues(); //TEMPORARY, maybe keep here as well

            if (streamTrackingStopwatch.IsRunning)
            {
                streamTrackingStopwatch.Stop();
                string streamTrackingDuration = streamTrackingStopwatch.Elapsed.ToString();
                ws.Cells[lastIntervalRow + 15, 13].NumberFormat = "HH:mm:ss";
                ws.Cells[lastIntervalRow + 15, 13] = streamTrackingDuration;
            }

            excel.Visible = false;
            wb.Save();
            wb.Close(true);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}