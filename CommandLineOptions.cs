using CommandLine;

namespace OutlookToTelegramNotifier
{
    class CommandLineOptions
    {
        [Option('t', "token", Required = true, HelpText = "Set Telegram bot token.")]
        public required string Token { get; set; }

        [Option('c', "chatId", Required = true, HelpText = "Set Telegram chat id.")]
        public required string ChatId { get; set; }

        [Option('i', "interval", Default = 5, Required = false, HelpText = "Set interval (in minutes) for checking messages.")]
        public int Interval { get; set; }

        [Option("sender-name-filter", Required = false, HelpText = "Set filter for the sender name field.")]
        public string SenderNameFilter { get; set; } = string.Empty;

        [Option("sender-emailaddress-filter", Required = false, HelpText = "Set filter for the sender email address field.")]
        public string SenderEmailAddressFilter { get; set; } = string.Empty;

        [Option("subject-filter", Required = false, HelpText = "Set filter for the subject field.")]
        public string SubjectFilter { get; set; } = string.Empty;

        [Option("important-only", Required = false, HelpText = "Set flag to process only messages marked as important.")]
        public bool ImportantOnly { get; set; }

        [Option('m', "mark-read", Required = false, HelpText = "Set flag to mark messages as read.")]
        public bool MarkRead { get; set; }

        [Option('a', "attach-as-pdf", Required = false, HelpText = "Set flag to send messages as PDF attachments.")]
        public bool AttachAsPDF { get; set; }

        [Option('d', "debug", Required = false, HelpText = "Set output to debug mode.")]
        public bool Debug { get; set; }
    }
}
