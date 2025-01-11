using CommandLine;
using Microsoft.Office.Interop.Outlook;
using System.Net.Http.Json;
using System.Text;
using Exception = System.Exception;
using Timer = System.Timers.Timer;

namespace OutlookToTelegramNotifier
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(options =>
                {
                    var timer = new Timer(options.Interval * 60 * 1000);                    
                    timer.Elapsed += async (sender, e) => 
                    {
                        try
                        {
                            Console.WriteLine($"{DateTime.Now.ToString()}: Checking messages.");

                            var nameSpace = (new Application()).GetNamespace("MAPI");

                            var filter = $"[Unread]=true AND [ReceivedTime] >= '{DateTime.Now.AddMinutes(-options.Interval):ddddd h:nn AMPM}'";                            

                            foreach (var mailItem in nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Items.Restrict(filter).OfType<MailItem>())
                            {
                                var applicable = true;

                                if (!string.IsNullOrEmpty(options.SenderNameFilter) && !mailItem.SenderName.Contains(options.SenderNameFilter))
                                {
                                    applicable = false;
                                }

                                if (!string.IsNullOrEmpty(options.SenderEmailAddressFilter) && !mailItem.SenderEmailAddress.Contains(options.SenderEmailAddressFilter))
                                {
                                    applicable = false;
                                }

                                if (!string.IsNullOrEmpty(options.SubjectFilter) && !mailItem.Subject.Contains(options.SubjectFilter))
                                {
                                    applicable = false;
                                }

                                if (options.ImportantOnly && mailItem.Importance != OlImportance.olImportanceHigh)
                                {
                                    applicable = false;
                                }

                                if (applicable)
                                {
                                    using var client = new HttpClient();
                                    var response = await client.PostAsJsonAsync($"https://api.telegram.org/bot{options.Token}/sendMessage",
                                        new
                                        {
                                            chat_id = options.ChatId,
                                            text = $"Unread message from \"{mailItem.SenderEmailAddress}\" with subject \"{mailItem.Subject}\"."
                                        });

                                    Console.WriteLine($"{DateTime.Now.ToString()}: {(response.IsSuccessStatusCode ? "Message sent to Telegram successfully." : "Failed to send message to Telegram.")}");
                                }

                                if (options.MarkRead)
                                {
                                    mailItem.UnRead = false;
                                }
                            }

                            nameSpace.Logoff();
                            nameSpace = null;
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
}