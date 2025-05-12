By default, whenever Outlook users are sending emails as mailboxes, the copy of the email is saved in their personal Sent folder and no copy is made in the Mailboxe's Sent folder.
This script is a simple solution to force Outlook to store emails sent through mailboxes both in personal account and in mailbox' sent folder.

**As any other Powershell file, you might need to unlock this one before opening through RMB > Properties > Unlock.**

The code is very straight-forward here.
1. We connect to Microsoft's Exchange througt ExchangeOnlineManagement moduel
2. Export all shared mailboxes to a CSV file
3. Afterwards, we apply the -MessageCopyForSentAsEnabled $True and -MessageCopyForSendOnBehalfEnabled $True commands to every single mailbox in the file.

That's it. Hope this saves you some time.
