using System;
using System.Threading.Tasks;
using Discord.Commands;
using Discord.WebSocket;
using Microsoft.Extensions.DependencyInjection;
using System.Reflection;
using Discord;

namespace TwitchChatDataBot
{
    class Program
    {
        public static TwitchChatBot bot;
        static void Main(string[] args)
        {
            bot = new TwitchChatBot();
            bot.Connect();

            //Start the discord bot
            new Program().RunDiscordBot().GetAwaiter().GetResult();

            Console.ReadLine();

            bot.Disconnect();
        }

        private DiscordSocketClient client;
        private CommandService commands;
        private IServiceProvider services;
        public async Task RunDiscordBot()
        {
            client = new DiscordSocketClient();
            commands = new CommandService();

            services = new ServiceCollection().AddSingleton(client).AddSingleton(commands).BuildServiceProvider();
            string botTOKEN = "NTIzNzUwNDk2MzYwNDY0Mzkz.DveERQ.0Lq2r3TD0jz9SQJ8IJeksCaFkpg";

            //event subscription
            client.Log += OnLog;

            await RegisterCommandAsync();

            await client.SetGameAsync("!twitchhelp");

            await client.LoginAsync(Discord.TokenType.Bot, botTOKEN);
            await client.StartAsync();
            //await Task.Delay(-1);
        }

        private Task OnLog(LogMessage arg)
        {
            Console.WriteLine(arg);
            return Task.CompletedTask;
        }

        public async Task RegisterCommandAsync()
        {
            client.MessageReceived += OnMessageReceived;
            await commands.AddModulesAsync(Assembly.GetEntryAssembly());
        }

        private async Task OnMessageReceived(SocketMessage arg)
        {
            var message = arg as SocketUserMessage;
            if (message == null || message.Author.IsBot)
                return;

            int argPos = 0;
            if (message.HasStringPrefix("!", ref argPos))
            {
                var context = new SocketCommandContext(client, message);
                var result = await commands.ExecuteAsync(context, argPos, services);

                if (!result.IsSuccess)
                    Console.WriteLine(result.ErrorReason);
            }
        }
    }
}
