---
title: 'Troubleshooting Office Scripts'
description: 'Debugging tips and techniques for Office Scripts, as well as help resources.'
ms.date: 07/20/2020
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

## Data limits

There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.

### Excel

Excel for the web has the following limitations when making calls to the workbook through a script:

- Requests and responses are limited to **5MB**.
- A range is limited to **five million cells**.

If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges. You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.

### Power Automate

When using Office Scripts with Power Automate, you're limited to **200 calls per day**. This limit resets at 12:00 AM UTC.

The Power Automate platform also other usage limitation, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).

## Help resources

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems. Often, you'll be able to find the solution to your problem through a quick Stack Overflow search. If not, ask your question and tag it with the "office-scripts" tag. Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.

If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository. Members of the product team will respond to issues and provide further assistance. Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.

If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.

## See also

- [Office Scripts in Excel on the web](../overview/excel.md)
- [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md)
- [Undo the effects of an Office Script](undo.md)
- [Improve the performance of your Office Scripts](../develop/web-client-performance.md)
