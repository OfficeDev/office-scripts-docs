---
title: Platform limits and requirements with Office Scripts
description: Resource limits and browser support for Office Scripts when used with Excel.
ms.date: 10/01/2022
ms.localizationpriority: medium
---

# Platform limits and requirements with Office Scripts

There are some platform limitations of which you should be aware when developing Office Scripts. This article details the browser support and data limits for Office Scripts for Excel.

## Data limits

There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.

### Excel

Excel for the web has the following limitations when making calls to the workbook through a script.

- Requests and responses are limited to **5MB**.
- A range is limited to **five million cells**.

If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges. For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample. You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) to target specific cells instead of large ranges.

### Power Automate

When using Office Scripts with Power Automate, each user is limited to **1,600 calls to the Run Script action per day**. This limit resets at 12:00 AM UTC.

The Power Automate platform also has usage limitations, which can be found in the following articles.

- [Limits and configuration in Power Automate](/power-automate/limits-and-config)
- [Known issues and limitations for the Excel Online (Business) connector](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> If you have a long-running script, be aware of the [120-second timeout for synchronous Power Automate operations](/power-automate/limits-and-config#timeout). You'll need to either [optimize your script](../develop/web-client-performance.md) or split your Excel automation into multiple scripts.

## Teams support

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## Third-party cookies for Excel on the web

Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web. Check your browser settings if the tab isn't being displayed. If you're using a private browser session, you may need to re-enable this setting each time.

> [!NOTE]
> Some browsers refer to this setting as "all cookies", instead of "third-party cookies".

### Instructions for adjusting cookie settings in popular browsers

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## Cross-platform API support

Some Office Scripts APIs may not be supported by Excel on Windows or Mac, especially older builds. These include newer APIs and APIs for web-only features. If a script contains unsupported APIs, the Code Editor displays a warning. If you try to run such a script, it won't run and, instead, the **Script Run Status** task pane displays a warning message that says, "This script currently must be run on Excel for the web. Open the workbook in the browser then try again, or contact the script owner for help."

## See also

- [Troubleshoot Office Scripts](troubleshooting.md)
- [Undo the effects of Office Scripts](undo.md)
- [Improve the performance of your Office Scripts](../develop/web-client-performance.md)
- [Scripting Fundamentals for Office Scripts in Excel](../develop/scripting-fundamentals.md)
