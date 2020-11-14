using System;
using System.IO;
using OfficeOpenXml;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;

namespace ClockBot2
{
    class Program
    {
        static ITelegramBotClient botClient;
        static FileInfo xlfile = new FileInfo("Schedule.xlsx");
        private ExcelPackage xlPack = new ExcelPackage(xlfile);
        
       
        static void Main(string[] args)
        {
            botClient = new TelegramBotClient("1462738929:AAHfFeF64pEUsvj7mIY2OUgK3AdeOgX_IJA");

            var me = botClient.GetMeAsync().Result;
            Console.WriteLine(
                $"Hello, World! I am user {me.Id} and my name is {me.FirstName}."
            );

            botClient.OnMessage += Bot_OnMessage;
            botClient.StartReceiving();

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();

            botClient.StopReceiving();
        }

        static async void Bot_OnMessage(object sender, MessageEventArgs e)
        {
            decimal time;
            if (e.Message.Text != null)
            {
                Console.WriteLine($"Received a text message in chat {e.Message.Chat.Id}.");
                string[] first = e.Message.Text.Split(' ');
                switch (first[0].ToUpper())
                {
                    case "СПРАВКА":
                    {
                        await botClient.SendTextMessageAsync(chatId: e.Message.Chat, text:
                            "__*Справка*__\n\n" +
                            @"*Справка*     \-      вызов справки" + "\n" +
                            @"*Пришел*      \-      запуск таймера" + "\n" +
                            @"*Ушел*        \-      останов таймера" + "\n" +
                            @"*Часы*        \-      записать кол\-во часов" + "\n",
                            parseMode: ParseMode.MarkdownV2,
                            replyToMessageId: e.Message.MessageId
                        );
                        break;
                    }
                    case "ПРИШЕЛ":
                    case "ПРИШЁЛ":
                    {
                        if (true)
                        {
                            await botClient.SendTextMessageAsync(chatId: e.Message.Chat, text:
                                "Таймер запущен для пользователя " + e.Message.Chat.FirstName + " " + e.Message.Chat.LastName);
                        };
                        break;
                    }
                    case "УШЕЛ":
                    case "УШЁЛ":
                    {
                        if (true)
                        {
                            await botClient.SendTextMessageAsync(chatId: e.Message.Chat, text:
                                "Таймер остановлен");
                        }

                        ;
                        break;
                    }
                    case "ЧАСЫ":
                    {
                        if ( first.Length == 1 )
                        {
                            await botClient.SendTextMessageAsync(chatId: e.Message.Chat, text:
                                "Не указано количество часов");
                        }
                        else
                        {
                            first[1] = first[1].Replace('.', ',');
                            if (!decimal.TryParse(first[1], out time))
                            {
                                await botClient.SendTextMessageAsync(chatId: e.Message.Chat, text:
                                            "Указано не корректное значение: не могу перевести в числовой формат");
                            }
                            else
                            {
                                    await botClient.SendTextMessageAsync(chatId: e.Message.Chat, text:
                                    "Часы успешно занесены в табель");
                            }
                            
                        }
                        break;
                    }
                }
                
            }
            
            
        }

        bool RecordTime(Chat chat)
        {
            return true;
        }
    }
}
