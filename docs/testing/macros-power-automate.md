---
title: 'Use macro files in Power Automate flows'
description: 'Learn how to use macro files or xlsm files in Power Automate flows.'
ms.date: 04/28/2021
localization_priority: Normal
---

# How to use macro files in Power Automate flows

[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.

However, macro files can't be selected in the file dropdown (see an example in the following screenshot).

:::image type="content" source="../images/no-xlsm.png" alt-text="The Power Automate Run script action showing no macro file selected. The error shown is 'File' is required.":::

One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="The Power Automate Run script action showing the macro file selected and no Run script error.":::

> [!NOTE]
> Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector. Be sure to test before deploying your solution.

## Other resources

[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).
