using System;
using System.Threading.Tasks;
using Discord.Commands;
using Microsoft.Office.Interop.Excel;

namespace TwitchChatDataBot.Discord_Modules
{
    public class GetIntervalData : ModuleBase<SocketCommandContext>
    {
        TwitchChatBot bot = TwitchChatDataBot.Program.bot;

        [Command("getintervaldata")]
        public async Task GetIntervalDataCommande(int intervalNum = -1, int streamNum = -1, string streamer = "")
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

            int streamCounter = 0;
            int numIntervals = 0;
            int rowNum = -1;
            int intervalRowNum = -1;

            if (streamNum < 0)
            {
                if (intervalNum < 0) //default to the newest stream
                {
                    foreach (Range r in ws.Rows)
                    {
                        if (ws.Cells[r.Row, 1].Value2 == null)
                            break;

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                        {
                            streamCounter++;
                            rowNum = r.Row;

                            if (streamNum > 0 && streamCounter <= streamNum)
                                numIntervals = 0;
                        }

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(double))
                            numIntervals++;
                    }
                }
                else //Target interval in newest stream
                {
                    foreach (Range r in ws.Rows)
                    {
                        if (ws.Cells[r.Row, 1].Value2 == null)
                            break;

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                        {
                            streamCounter++;
                            rowNum = r.Row;

                            numIntervals = 0;
                        }

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(double))
                        {
                            numIntervals++;
                            if (numIntervals == intervalNum)
                                intervalRowNum = r.Row;
                        }
                    }
                }
            }
            else //streamNum is greater than 0
            {
                if(intervalNum < 0)
                {
                    foreach (Range r in ws.Rows) {
                        if (streamCounter > streamNum || ws.Cells[r.Row, 1].Value2 == null)
                            break;

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                        {
                            streamCounter++;
                            
                            if (streamCounter <= streamNum)
                            {
                                rowNum = r.Row;
                                numIntervals = 0;
                            }
                        }

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(double))
                            numIntervals++;
                    }
                }
                else //streamNum not -1 AND intervalNum not -1
                {
                    foreach (Range r in ws.Rows)
                    {
                        if (streamCounter > streamNum || ws.Cells[r.Row, 1].Value2 == null)
                            break;

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                        {
                            streamCounter++;

                            if (streamCounter <= streamNum)
                            {
                                rowNum = r.Row;
                                numIntervals = 0;
                            }
                        }

                        if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(double))
                        {
                            numIntervals++;
                            if (numIntervals == intervalNum)
                                intervalRowNum = r.Row;
                        }
                    }
                }
            }

            int targetCellRow = -1;
            if (intervalNum < 0)
                targetCellRow = rowNum + numIntervals;
            else
                targetCellRow = intervalRowNum;

            targetCellRow -= 3;

            if(numIntervals == 0)
            {
                await ReplyAsync("Currently no data, please try again in a few minutes.");
            }
            else
            {   
                await ReplyAsync(string.Format(
                "**Streamer:** {7}\n" +
                "**Interval:** {0}\n" +
                "**Average Viewers:** {1}\n" +
                "**Messages Sent:** {2}\n" +
                "**Emotes Sent:** {3}\n" +
                "**Most Popular Emotes:** {4}\n" +
                "**Followers Gained:** {5}\n" +
                "**Subscribers Gained:** {6}",
                ws.Cells[targetCellRow, 1].Value2, ws.Cells[targetCellRow, 2].Value2, ws.Cells[targetCellRow, 3].Value2, ws.Cells[targetCellRow, 4].Value2, 
                ws.Cells[targetCellRow, 5].Value2, ws.Cells[targetCellRow, 6].Value2, ws.Cells[targetCellRow, 7].Value2, ws.Name));
                
            }
        }
    }
}
