using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Discord.Commands;

namespace TwitchChatDataBot.Discord_Modules
{
    public class HelpCommand : ModuleBase<SocketCommandContext>
    {
        string[] commands = new string[] {
            "**!getstreamers** | Returns a list of streamers that have been tracked",
            "**!getactivestream** | Returns current stream",
            "**!getintervals[streamNum = -1]** | Returns number data intervals in a stream",
            "**!getintervaldata[intervalNum = -1][streamNum = -1][streamer = '']** | Return data from a specific interval in a stream collection",
            "**!gettotaldata[streamNum = -1][streamer = '']** | Returns data found in the total column" };

        [Command("twitchhelp")]
        public async Task DisplayCommands()
        {
            await ReplyAsync("**Note: Intervals are measured in 5 minute timespans.**\n\nCommand List:");
            foreach (string s in commands)
                await ReplyAsync(s + "\n");
        }
    }
}
