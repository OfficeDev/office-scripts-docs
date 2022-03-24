---
title: Use macro-enabled files in Power Automate flows
description: Learn how to use macro-enabled files, or .xlsm files, in Power Automate flows.
ms.date: 09/01/2021
ms.localizationpriority: medium
---

# How to use macro-enabled files in Power Automate flows

You can integrate your .xlsm files with a Power Automate flow. This lets you start converting your existing automation solutions to web-based formats. Please note that the macros contained in the .xslm files cannot be run through Power Automate. Only Office Scripts are enabled there.

The [Excel Online (Business) connector](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) in [Power Automate](https://flow.microsoft.com/) is typically limited to files in the Microsoft Excel Open XML Spreadsheet (.xlsx) format. Its file browser only lets you select .xlsx files. However, macro-enabled files are compatible with the connector's **Run script** action if the file metadata is used.

In your flow, use the **Get File Metadata** action from either the [OneDrive for Business](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) or [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) connectors. The **Run script** action accepts this metadata as a valid file. Use the *ID* dynamic content returned from the **Get file metadata** action as the "File" argument when running the script. The following screenshot shows a flow providing the metadata for a file called "Test Macro File.xlsm" to a **Run script** action.

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="A flow with a Get file metadata action passing the metadata of a macro file to a Run script action.":::

> [!WARNING]
> Some .xlsm files, especially those with ActiveX or Form controls, may not work in the Excel online connector. Be sure to test before deploying your solution.

## Other resources

[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).
