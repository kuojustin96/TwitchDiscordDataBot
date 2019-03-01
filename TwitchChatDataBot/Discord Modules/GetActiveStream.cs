using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Discord.Commands;
using Microsoft.Office.Interop.Excel;

namespace TwitchChatDataBot.Discord_Modules
{
    public class GetActiveStream : ModuleBase<SocketCommandContext>
    {
        TwitchChatBot bot = TwitchChatDataBot.Program.bot;

        [Command("getactivestream")]
        public async Task GetStreamAsync()
        {
            Worksheet ws = bot.excel.ActiveSheet;
            string wsName = ws.Name;
            //int intervalRow = bot.lastIntervalRow;
            int rowNum = -1;
            foreach (Range r in ws.Rows)
            {
                if (ws.Cells[r.Row, 1].Value2 == null)
                    break;

                if (ws.Cells[r.Row, 1].Value2.GetType() == typeof(string) && ws.Cells[r.Row, 1].Value2.Contains("Interval"))
                {
                    rowNum = r.Row;
                }
            }

            await ReplyAsync(string.Format("Streamer being tracked: {0}\n{1}", ws.Name, ws.Cells[rowNum - 1, 1].Value2));
        }
    }
}