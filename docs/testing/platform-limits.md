---
title: 'Platform limits and requirements with Office Scripts'
description: 'Resource limits and browser support for Office Scripts when used with Excel on the web'
ms.date: 03/12/2021
localization_priority: Normal
---

# Platform limits and requirements with Office Scripts

There are some platform limitations of which you should be aware when developing Office Scripts. This article details the browser support and data limits for Office Scripts for Excel on the web.

## Browser support

Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11). Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11. If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### Third-party cookies

Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web. Check your browser settings if the tab isn't being displayed. If you're using a private browser session, you may need to re-enable this setting each time.

> [!NOTE]
> Some browsers refer to this setting as "all cookies", instead of "third-party cookies".

#### Instructions for adjusting cookie settings in popular browsers

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## Data limits

There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.

### Excel

Excel for the web has the following limitations when making calls to the workbook through a script:

- Requests and responses are limited to **5MB**.
- A range is limited to **five million cells**.

If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges. You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.

### Power Automate

When using Office Scripts with Power Automate, each user is limited to **400 executions of the Run Script action per day**. This limit resets at 12:00 AM UTC.

The Power Automate platform also has usage limitations, which can be found in the following articles:

- [Limits and configuration in Power Automate](/power-automate/limits-and-config)
- [Known issues and limitations for the Excel Online (Business) connector](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## See also

- [Troubleshooting Office Scripts](troubleshooting.md)
- [Undo the effects of an Office Script](undo.md)
- [Improve the performance of your Office Scripts](../develop/web-client-performance.md)
- [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md)
