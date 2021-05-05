---
title: 'Office Scripts file storage and ownership'
description: 'Information about how Office Scripts are stored in Microsoft OneDrive and transferred between owners.'
ms.date: 11/13/2020
localization_priority: Normal
---

# Office Scripts file storage and ownership

Office Scripts are stored as **.osts** files in your Microsoft OneDrive. This allows your scripts to exist outside any particular workbook. Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.

## File storage

You Office Scripts are stored in your OneDrive. The **.osts** files are found in the **/Documents/Office Scripts/** folder. Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.

Scripts that are shared with one of your workbooks remain in the script creator's OneDrive. They are not copied to any of your local or OneDrive folders when you run the shared script in Excel. The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive. Changes to the copy don't affect the original script.

### Script folders

Adding folders to your OneDrive helps keep your scripts organized. Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor. Please note that these folders cannot be created or deleted by using the Code Editor. Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.

:::image type="content" source="../images/script-folders.png" alt-text="The New Script dialog in Code Editor showing scripts contained in folders, as displayed in task pane":::

## File ownership and retention

Office Scripts are stored in a user's OneDrive. They follow the retention and deletion policies specified by Microsoft OneDrive. To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).

## See also

- [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Office Scripts settings in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Undo the effects of an Office Script](../testing/undo.md)
