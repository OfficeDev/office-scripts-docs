---
title: 'Use macro files in Power Automate flows'
description: 'Learn how to use macro files or xlsm files in Power Automate flows.'
ms.date: 09/01/2021
localization_priority: Normal
---

# How to use macro files in Power Automate flows

The [Excel Online (Business) connector](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) in [Power Automate](https://flow.microsoft.com/) typically only works with files in the Microsoft Excel Open XML Spreadsheet (.xlsx) format. The file browser limits your selection to .xlsx files inside the connector. However, macro files are compatible with the connector's **Run script** action if the file metadata is used.

In your flow, use the **Get File Metadata** action from either the [OneDrive for Business](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) or [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) connectors. The **Run script** action accepts this metadata as a valid file. Use the *ID* dynamic content returned from the **Get file metadata** action as the "File" argument when running the script. The following screenshot shows a flow providing the metadata for a file called "Test Macro File.xlsm" to a **Run script** action.

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="A flow with a Get file metadata action passing the metadata of a macro file to a Run script action.":::

> [!WARNING]
> Some .xlsm files, especially the ones with ActiveX or Form controls, may not work in the Excel online connector. Be sure to test before deploying your solution.

## Other resources

[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).
