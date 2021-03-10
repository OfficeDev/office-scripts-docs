---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Using Macro files or xlsm files in Power Automate flows

[Power Automation flows](https://us.flow.microsoft.com/) provides [Excel connectors](https://us.flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with rest of your organizational data/apps such as Teams, Outlook, SharePoint, etc. 

However, one of the limitation is that macro files cannot be selected in the file drop-down. See the screen shot below. 
'
One way to get around this issue is by including "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as showin in the second screenshot. 

_Note_: Some XLSM (esp. the ones with ActiveX/Form controls) may not work in Excel online connector. Be sure to test before deploying your solution. 

Hope this helps! 

_No XLSM  in Excel Online connector_ 

![No xlsm in Run Script action](no%20xlsm.png)


_Use XLSM Macro file in Excel Online Run Script actinon_

![xlsm in Run Script action](xlsm%20in%20pa.png)


[![Use XLSM in Run Script action](v_xlsm.png)](https://youtu.be/o-H9BbywJQQ)
