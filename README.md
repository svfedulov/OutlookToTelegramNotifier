# Outlook to Telegram notifier program

The program is designed to monitor new messages in Microsoft Outlook and send notifications about them to Telegram. If you are unable or prefer not to connect directly to your mailbox, this solution ensures you stay informed at all times.

## Features

- Unread email detection in the Microsoft Outlook Inbox folder.

The program connects not to the mail server but directly to the mail client installed on the workstation. This architecture has both advantages and disadvantages:

Advantages:
	Corporate mail servers often impose restrictions on direct connections (e.g., network, client, etc.). In this case, you connect from a corporate device, often using a VPN connection and a standard mail client.
	There is no need to enter authentication credentials into a third-party application.

Disadvantages:
	The program operates only on Microsoft Windows and requires Microsoft Outlook to be installed.
	The device must remain powered on and connected to the mail server.

- Notification delivery to Telegram messenger.
- Configurable message check Intervals.
- Message filtering options: filter by sender name, email address, subject, or message importance.
- Mark messages as read.
- Send messages as PDF attachments to Telegram.

Initially, attachments were sent in MSG format, but it turned out that the mobile version of Microsoft Outlook does not support this format, which is standard for the desktop client.

## Getting started

## To-do list

- Get rid of "OutlookToTelegramNotifier.Program[0]" in console.

## Known Issues

- For some messages the SenderEmailAddress field looks like "/O={Organization}/OU={Organizational unit}/CN={User}. At the moment I don't have a solution to this problem. The user's real address is not among the message fields.

## Change log

### 1.0.3.0

- Fixed: Issue with both angle brackets in the subject field.

### 1.0.2.0

- Added: Added and implemented debug mode (instead of verbose mode).
- Added: Added flag to send messages as PDF attachments.

### 1.0.1.0

- Fixed: Fixed the error measuring the time of the previous check.
- Fixed: Output texts in Unicode format to the console.
- Added: Added filter by sender name.
- Added: Added filter by sender email address.
- Added: Added filter by subject.
- Added: Added flag to process only messages marked as important.
- Added: Added flag to mark messages as read.
- Added: Added verbose mode (not implemented).

### 1.0.0.0

- Initial version with basic function.