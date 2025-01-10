using CommandLine;
using Microsoft.Office.Interop.Outlook;
using System.Net.Http.Json;
using Exception = System.Exception;
using Timer = System.Timers.Timer;

class Options
{
    [Option('t', "token", Required = true, HelpText = "Set Telegram bot token.")]
    public required string Token { get; set; }

    [Option('c', "chatId", Required = true, HelpText = "Set Telegram chat id.")]
    public required string ChatId { get; set; }

    [Option('i', "interval", Default = 5, Required = false, HelpText = "Set interval in minutes for checking messages.")]
    public int Interval { get; set; }
}

class Program
{
    static void Main(string[] args)
    {
        Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(options =>
            {
                DateTime ago = DateTime.Now.AddMinutes(-options.Interval);

                Timer timer = new Timer(options.Interval * 60 * 1000);
                timer.Elapsed += async (sender, e) => {
                    try
                    {
                        Console.WriteLine($"{DateTime.Now.ToString()}: Checking messages...");

                        Application? application = new Application();
                        NameSpace? nameSpace = application.GetNamespace("MAPI");
                        MAPIFolder? mAPIFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                        Items? items = mAPIFolder.Items.Restrict("[Unread]=true");

                        foreach (MailItem mailItem in items.OfType<MailItem>())
                        {
                            if (mailItem.ReceivedTime >= ago)
                            {
                                var notification = $"New unread message from {mailItem.SenderEmailAddress} with subject {mailItem.Subject}.";

                                Console.WriteLine($"{DateTime.Now.ToString()}: {notification}");

                                using (HttpClient client = new HttpClient())
                                {
                                    HttpResponseMessage response = await client.PostAsJsonAsync($"https://api.telegram.org/bot{options.Token}/sendMessage",
                                        new
                                        {
                                            chat_id = options.ChatId,
                                            text = notification
                                        });

                                    Console.WriteLine($"{DateTime.Now.ToString()}: {(response.IsSuccessStatusCode ? "Message sent to Telegram successfully." : "Failed to send message to Telegram.")}");
                                }
                            }
                        }

                        nameSpace.Logoff();

                        items = null;
                        mAPIFolder = null;
                        nameSpace = null;
                        application = null;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{DateTime.Now.ToString()}: Error: {ex.Message}");
                    }
                };
                timer.AutoReset = true;
                timer.Enabled = true;

                Console.WriteLine($"{DateTime.Now.ToString()}: Service started. Press Enter to exit.");
                Console.ReadLine();
                Console.WriteLine($"{DateTime.Now.ToString()}: Service stopped.");
            });
    }
}