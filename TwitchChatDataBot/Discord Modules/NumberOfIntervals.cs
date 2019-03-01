using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Discord.Commands;
using Microsoft.Office.Interop.Excel;

namespace TwitchChatDataBot.Discord_Modules
{
    public class NumberOfIntervals : ModuleBase<SocketCommandContext>
    {
        TwitchChatBot bot = TwitchChatDataBot.Program.bot;

        [Command("getintervals")]
        public async Task GetNumIntervals(int streamNum = -1, string streamer = "")
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
                await ReplyAsync("Stream invalid, please try again.");
                return;
            }

            int streamCounter = 0;
            int numIntervals = 0;

            if (streamNum < 0) //default to the newest stream
            {
                foreach (Range r in ws.Rows)
                {
                    if (ws.Cells[r.Row, 1].Value2 == null)
                        break;

                    if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                    {
                        streamCounter++;
                        numIntervals = 0;
                    }

                    if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(double))
                        numIntervals++;
                }
            }
            else
            {
                foreach (Range r in ws.Rows)
                {
                    if (streamCounter > streamNum || ws.Cells[r.Row, 1].Value2 == null)
                    {
                        streamCounter -= 1;
                        break;
                    }

                    if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                    {
                        streamCounter++;

                        if(streamCounter <= streamNum)
                            numIntervals = 0;
                    }

                    if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(double))
                        numIntervals++;
                }
            }

            await ReplyAsync(string.Format("Stream {0} by {1} has {2} intervals", streamCounter, ws.Name, numIntervals));
        }
    }
}
