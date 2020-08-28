---
title: 'Troubleshooting Office Scripts'
description: 'Debugging tips and techniques for Office Scripts, as well as help resources.'
ms.date: 07/23/2020
localization_priority: Normal
---

# Troubleshooting Office Scripts

As you develop Office Scripts, you may make mistakes. It's okay. We have tools that help find the problems and get your scripts working perfectly.

## Console logs

Sometimes while troubleshooting, you'll want to print messages to the screen. These can show you the current value of variables or which code paths are being triggered. To do this, log text to the console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Strings passed to`console.log` will be displayed in the Code Editor's logging console. To turn on the console, press the **Ellipses** button and select **Logs...**

Logs do not affect the workbook.

## Error messages

When your Excel Script encounters a problem running, it produces an error. You'll see a prompt pop-up asking if you want to **View Logs**. Press that button to open the console and display any errors.

## Automate tab not appearing

The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel for the web.

1. [Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).
1. [Have your admin enable the feature](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. [Check that your browser is supported](platform-limits.md#browser-support).
1. [Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).

## Help resources

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems. Often, you'll be able to find the solution to your problem through a quick Stack Overflow search. If not, ask your question and tag it with the "office-scripts" tag. Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.

If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository. Members of the product team will respond to issues and provide further assistance. Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.

If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.

## See also

- [Office Scripts in Excel on the web](../overview/excel.md)
- [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md)
- [Platform Limits with Office Scripts](platform-limits.md)
- [Improve the performance of your Office Scripts](../develop/web-client-performance.md)
- [Undo the effects of an Office Script](undo.md)
