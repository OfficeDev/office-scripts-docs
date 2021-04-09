---
title: 'Use macro files in Power Automate flows'
description: 'Learn how to use macro files or xlsm files in Power Automate flows.'
ms.date: 03/18/2021
localization_priority: Normal
---

# How to use macro files in Power Automate flows

[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.

However, macro files can't be selected in the file dropdown (see an example in the following screenshot).

:::image type="content" source="../images/no-xlsm.png" alt-text="No xlsm in Run Script action":::

One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="xlsm in Run Script action":::

> [!NOTE]
> Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector. Be sure to test before deploying your solution.

[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")
