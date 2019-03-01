using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Discord.Commands;
using Microsoft.Office.Interop.Excel;

namespace TwitchChatDataBot.Discord_Modules
{
    public class GetStreamers : ModuleBase<SocketCommandContext>
    {
        TwitchChatBot bot = TwitchChatDataBot.Program.bot;

        [Command("getstreamers")]
        public async Task GetStreamersCommand()
        {
            await ReplyAsync("**List of Streamers Tracked:**\n");
            foreach (Worksheet ws in bot.excel.Sheets)
                if(ws.Name != "Template Worksheet")
                    await ReplyAsync("**" + ws.Name + "**\n");
        }
    }
}
