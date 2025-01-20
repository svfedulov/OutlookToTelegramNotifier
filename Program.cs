using CommandLine;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Telegram.Bot;
using Telegram.Bot.Types;

namespace OutlookToTelegramNotifier
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            Parser.Default.ParseArguments<CommandLineOptions>(args).WithParsed<CommandLineOptions>(options => 
            {
                using ILoggerFactory loggerFactory = LoggerFactory.Create(builder =>
                {
                    builder.AddSimpleConsole(options =>
                    {
                        options.IncludeScopes = false;
                        options.SingleLine = true;
                        options.TimestampFormat = $"{CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern} {CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern} "; 
                    });

                    builder.SetMinimumLevel(options.Debug ? LogLevel.Debug : LogLevel.Information);
                });

                ILogger<Program> logger = loggerFactory.CreateLogger<Program>();

                var timer = new System.Timers.Timer(options.Interval * 60 * 1000);                
                timer.Elapsed += async (sender, e) => 
                {
                    try
                    {
                        logger.LogInformation("Checking unread messages.");

                        var nameSpace = (new Application()).GetNamespace("MAPI");

                        var mailItems = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Items.Restrict("[Unread]=true").OfType<MailItem>();

                        logger.LogInformation("Found {Count} unread message(s).", mailItems.Count());

                        foreach (var mailItem in mailItems)
                        {
                            logger.LogInformation("Processing message with SenderName: {SenderName}, SenderEmailAddress: {SenderEmailAddress}, Subject: {Subject}, Importance: {Importance}.", 
                                mailItem.SenderName, 
                                mailItem.SenderEmailAddress, 
                                mailItem.Subject,
                                new Dictionary<int, string>() { { 0, "Low" }, { 1, "Normal" }, { 2, "High" } }.GetValueOrDefault((int)mailItem.Importance, ""));

                            if (
                                mailItem.ReceivedTime < DateTime.Now.AddMinutes(-options.Interval) ||
                                !string.IsNullOrEmpty(options.SenderNameFilter) && !mailItem.SenderName.Contains(options.SenderNameFilter) ||
                                !string.IsNullOrEmpty(options.SenderEmailAddressFilter) && !mailItem.SenderEmailAddress.Contains(options.SenderEmailAddressFilter) ||
                                !string.IsNullOrEmpty(options.SubjectFilter) && !mailItem.Subject.Contains(options.SubjectFilter) ||
                                options.ImportantOnly && mailItem.Importance != OlImportance.olImportanceHigh
                                )
                            {
                                logger.LogDebug("The message does not match the specified filters.");
                            }
                            else
                            {
                                logger.LogDebug("The message matches the specified filters.");

                                var telegramBotClient = new TelegramBotClient(options.Token);

                                var me = await telegramBotClient.GetMe();
                                logger.LogDebug("Connected to Telegram bot with id: {Id} and username: {Username}.", me.Id, me.Username);

                                if (options.AttachAsPDF)
                                {
                                    logger.LogDebug("Send notification with an attached message in PDF format.");

                                    var fileName = $"{Path.GetTempPath()}{mailItem.EntryID}.pdf";

                                    var inspector = mailItem.GetInspector;

                                    if (inspector.IsWordMail() && inspector.EditorType == OlEditorType.olEditorWord)
                                    {
                                        var wordEditor = inspector.WordEditor;
                                        wordEditor.ExportAsFixedFormat(fileName, 17);

                                        Marshal.ReleaseComObject(wordEditor);
                                        wordEditor = null;
                                    }

                                    try
                                    {
                                        await using Stream stream = File.OpenRead(fileName);

                                        var message = await telegramBotClient.SendDocument(options.ChatId,
                                            document: InputFile.FromStream(stream, $"{mailItem.EntryID}.pdf"),
                                            caption: $"<b>New unread message</b>\n<b>SenderName</b>: {mailItem.SenderName}\n<b>SenderEmailAddress:</b> {mailItem.SenderEmailAddress}\n" +
                                            $"<b>Subject:</b> {mailItem.Subject}\n<b>Importance:</b> {new Dictionary<int, string>() { { 0, "Low" }, { 1, "Normal" }, { 2, "High" } }.GetValueOrDefault((int)mailItem.Importance, "")}",
                                            parseMode: Telegram.Bot.Types.Enums.ParseMode.Html);

                                        logger.LogDebug("Sent message with id: {Id} to Username: {Username}.", message.Id, message.Chat.Username);

                                        File.Delete(fileName);
                                    }
                                    catch
                                    {

                                    }
                                }
                                else
                                {
                                    logger.LogDebug("Send a notification without attachment.");

                                    var message = await telegramBotClient.SendMessage(options.ChatId, 
                                        $"<b>New unread message</b>\n<b>SenderName:</b> {mailItem.SenderName}\n<b>SenderEmailAddress:</b> {mailItem.SenderEmailAddress}\n" +
                                        $"<b>Subject:</b> {mailItem.Subject}\n<b>Importance:</b> {new Dictionary<int, string>() { { 0, "Low" }, { 1, "Normal" }, { 2, "High" } }.GetValueOrDefault((int)mailItem.Importance, "")}",
                                        parseMode: Telegram.Bot.Types.Enums.ParseMode.Html);

                                    logger.LogDebug("Sent message with id: {Id} to Username: {Username}.", message.Id, message.Chat.Username);
                                } 
                            }

                            if (options.MarkRead)
                            {
                                logger.LogDebug("Mark message as read.");

                                mailItem.UnRead = false;
                            }
                        }

                        nameSpace.Logoff();
                        //Marshal.ReleaseComObject(nameSpace);
                        nameSpace = null;
                    }
                    catch (COMException ex) 
                    {
                        if (ex.HResult == -2147221164)
                        {
                            logger.LogError("Outlook is not installed on the system.");
                        }
                        else if (ex.HResult == -2079195127)
                        {
                            logger.LogError("Outlook is not connected.");
                        }
                        else if (ex.HResult == -2147467260)
                        {
                            logger.LogError("Operation aborted by the user.");
                        }
                        else
                        {
                            logger.LogError("Exception caught: {Message}", ex.ToString());
                        }
                    }
                    catch (System.Exception ex)
                    {
                        logger.LogError("Exception caught: {Message}", ex.ToString());
                    }
                };
                timer.AutoReset = true;
                timer.Enabled = true;

                Console.WriteLine($"{Process.GetCurrentProcess().MainModule?.FileVersionInfo.ProductName} {Process.GetCurrentProcess().MainModule?.FileVersionInfo.ProductVersion}");
                Console.WriteLine(Process.GetCurrentProcess().MainModule?.FileVersionInfo.LegalCopyright);
                Console.WriteLine();

                logger.LogInformation("Program started. Press Enter to exit.");
                
                Console.ReadLine();
                
                logger.LogInformation("Program stopped.");                    
            });
        }
    }
}