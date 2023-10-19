---
title: Use macro-enabled files in Power Automate flows
description: Learn how to use macro-enabled files, or .xlsm files, in Power Automate flows.
ms.date: 10/19/2023
ms.localizationpriority: medium
---

# How to use macro-enabled files in Power Automate flows

[Power Automate](https://make.powerautomate.com/) flows support .xlsm files in the [Excel Online (Business) connector](https://make.powerautomate.com/connectors/shared_excelonlinebusiness/excel-online-business/).

> [!IMPORTANT]
> The macros contained in the .xlsm files cannot be run through Power Automate. Only Office Scripts are enabled there.

The **Run script** action lets you select .xlsm files. Other connector actions require you use the file ID. Get this ID with the **Get File Metadata** action from either the [OneDrive for Business](https://make.powerautomate.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) or [SharePoint](https://make.powerautomate.com/connectors/shared_sharepointonline/sharepoint/) connectors. Use the *ID* dynamic content returned from the **Get file metadata** action as the "File" argument for Excel connector actions.

> [!WARNING]
> Some .xlsm files, especially those with ActiveX or Form controls, may not work in the Excel online connector. Be sure to test before deploying your solution.

## See also

- [Excel Online (Business) connector reference](/connectors/excelonlinebusiness/)
- [Run Office Scripts with Power Automate](power-automate-integration.md)