using System;
using System.Threading.Tasks;
using Discord.Commands;
using Microsoft.Office.Interop.Excel;

namespace TwitchChatDataBot.Discord_Modules
{
    public class GetTotalData : ModuleBase<SocketCommandContext>
    {
        TwitchChatBot bot = TwitchChatDataBot.Program.bot;

        [Command("gettotaldata")]
        public async Task GetTotalDataCommand(int streamNum = -1, string streamer = "")
        {
            Worksheet ws = null;
            if (streamer == "")
            {
                Console.WriteLine("No streamer given, using active sheet");
                ws = bot.excel.ActiveSheet;
            }
            else
            {
                Console.WriteLine("Searching for " + streamer);
                bool foundSheet = false;
                foreach (Worksheet sheet in bot.wb.Sheets)
                {
                    if (sheet.Name == streamer)
                    {
                        ws = sheet;
                        foundSheet = true;
                        break;
                    }
                }

                if (!foundSheet)
                {
                    await ReplyAsync("Could not find specified worksheet, please try again.");
                    return;
                }
            }

            if (streamNum > ws.Cells[3, 16].Value2 || streamNum == 0)
            {
                await ReplyAsync("Stream number invalid, please try again.");
                return;
            }


            int rowNum = -1;
            int streamCounter = 0;
            int targetStream = -1;
            if (streamNum < 0) //defaults to newest stream
            {
                foreach (Range r in ws.Rows)
                {
                    if (ws.Cells[r.Row, 1].Value2 == null)
                        break;

                    if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                        rowNum = r.Row;
                }

                targetStream = Convert.ToInt32(ws.Cells[3, 16].Value2);
            }
            else //Get data from a specific stream number
            {
                foreach (Range r in ws.Rows)
                {
                    if (streamCounter > streamNum || ws.Cells[r.Row, 1].Value2 == null)
                        break;

                    if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                    {
                        streamCounter++;

                        if (streamCounter <= streamNum)
                            rowNum = r.Row;
                    }
                }

                targetStream = streamCounter - 1;
            }

            //Make sure there's actually data to read
            if (ws.Cells[rowNum + 2, 13].Value2 == null)
            {
                await ReplyAsync("Currently no data available, please try again in a few minutes.");
                return;
            }

            await ReplyAsync(string.Format(
                "**Total Data From Stream {0} by {1}**\n" +
                "**Total Viewers:** {2}\n" +
                "**Most Concurrent Viewers:** {3}\n" +
                "**Average Viewers:** {4}\n" +
                "**Followers Gained:** {5}\n" +
                "**Avg Followers Gained Per Interval:** {6}", 
                targetStream, ws.Name, ws.Cells[rowNum + 2, 13].Value2, ws.Cells[rowNum + 3, 13].Value2,
                ws.Cells[rowNum + 4, 13].Value2, ws.Cells[rowNum + 5, 13].Value2, ws.Cells[rowNum + 12, 13].Value2));
            await ReplyAsync(string.Format(
                "**Subscribers Gained:** {0}\n" +
                "**Avg Subs Gained Per Interval:** {1}\n" +
                "**Total Emotes Used:** {2}\n" +
                "**Avg Emotes Used Per Interval:** {3}\n" +
                "**Most Popular Emotes:** {4}\n" +
                "**Total Messages Sent:** {5}\n" +
                "**Avg Messages Sent Per Interval:** {6}\n",
                ws.Cells[rowNum + 6, 13].Value2, ws.Cells[rowNum + 13, 13].Value2, ws.Cells[rowNum + 7, 13].Value2, ws.Cells[rowNum + 11, 13].Value2,
                ws.Cells[rowNum + 8, 13].Value2, ws.Cells[rowNum + 9, 13].Value2, ws.Cells[rowNum + 10, 13].Value2));
            string streamUptime = DateTime.FromOADate(ws.Cells[rowNum + 14, 13].Value2).ToString("HH:mm:ss");
            string streamTracking = DateTime.FromOADate(ws.Cells[rowNum + 15, 13].Value2).ToString("HH:mm:ss");
            await ReplyAsync(string.Format("**Stream Uptime:** {0}\n**Stream Tracking Duration:** {1}",
                streamUptime, streamTracking));
        }
    }
}
