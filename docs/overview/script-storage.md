---
title: 'Office Scripts file storage and ownership'
description: 'Information about how Office Scripts are stored in Microsoft OneDrive and transferred between owners.'
ms.date: 05/17/2021
localization_priority: Normal
---

# Office Scripts file storage and ownership

Office Scripts are stored as **.osts** files in your Microsoft OneDrive. They are stored separately from a workbook. To give others access, [share the script with an Excel workbook](excel.md#sharing-scripts). This means you're linking the script with the file, not attaching it. Whoever has access to the Excel file will also be able to view, run, or make a copy of the script.

Unless you share your scripts, no one else can access them. Your OneDrive settings control the shared access and permissions for all script **.osts** files, independent of any Excel settings. Scripts can't be linked from a local disk or custom cloud locations. Office Scripts only recognizes and runs a script if it's in your OneDrive folder or shared with the workbook.

## File storage

You Office Scripts are stored in your OneDrive. The **.osts** files are found in the **/Documents/Office Scripts/** folder. Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.

Scripts that are shared with one of your workbooks remain in the script creator's OneDrive. They are not copied to any of your local or OneDrive folders when you run the shared script in Excel. The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive. Changes to the copy don't affect the original script.

## File ownership and retention

Office Scripts are stored in a user's OneDrive. They follow the retention and deletion policies specified by Microsoft OneDrive. To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).

During editing, files are temporarily stored in the browser. You must save the script before closing the Excel window to save it to the OneDrive location. Don't forget to save the file after edits, or else those edits will only be in the browser's version of the file.

## See also

- [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Office Scripts settings in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Undo the effects of an Office Script](../testing/undo.md)
