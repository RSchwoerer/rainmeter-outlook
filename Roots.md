In the measure that fetches the email folders you can use the Root parameter to specify your email folder locations. Use a pipe | character to specify multiple roots. Small example:

```
[MeasureUnreadEmails]
Measure=Plugin
Plugin=Plugins\OutlookPlugin.dll
Resource=MAPIFolder
Root=\\MyAccount1|\\MyAccount2\Inbox
Filter=%HasUnreadItems
```

This will include all emails from the folder MyAccount1, including junk mail and the recycle bin, and the inbox of the folder MyAccount2.

If you are using IMAP, the default folder name is the email address. The name of the inbox depends on your locale.