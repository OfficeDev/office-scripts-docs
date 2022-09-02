---
title: Office Scripts file storage and ownership
description: Information about how Office Scripts are stored in Microsoft OneDrive and transferred between owners.
ms.date: 08/31/2022
ms.localizationpriority: medium
---

# Office Scripts file storage and ownership

Office Scripts are stored as **.osts** files in your Microsoft OneDrive or a SharePoint folder. They are stored separately from a workbook. To give users who are outside of the SharePoint site access to the script, [share the script with an Excel workbook](excel.md#share-office-scripts). This means you're linking the script with the file, not attaching it. Whoever has access to the Excel file will also be able to view, run, or make a copy of the script.

Excel only recognizes and runs a script if it's in your OneDrive folder, a Sharepoint folder, or shared with the workbook.

## OneDrive

The default behavior is that Office Scripts are stored in your OneDrive. The **.osts** files are found in the **/Documents/Office Scripts/** folder. Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.

Scripts that are shared with one of your workbooks remain in the script creator's OneDrive. They are not copied to any of your local or OneDrive folders when you run the shared script in Excel. The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive. Changes to the copy don't affect the original script.

Unless you share your personal scripts, no one else can access them. Your OneDrive settings control the shared access and permissions for all script **.osts** files, independent of any Excel settings. Scripts can't be linked from a local disk or custom cloud locations.

## SharePoint

Office Scripts that are saved to a SharePoint site are owned by your team. You and members of your organization with the appropriate access can run and edit scripts from SharePoint. You'll also see these scripts appear in the **Automate** tab's Script Gallery.

To load a script from SharePoint, go to **All scripts** and select **View more scripts** at the bottom of the list. This brings up a file picker where you can choose **.osts** files from any SharePoint site to which you have access. Note that scripts from SharePoint that you've already opened will be displayed in the list of recent scripts.

To save a script to SharePoint, go to the **More options (…)** menu and select **Save as**. This opens a file picker where you can select folders in your SharePoint site. Saving to a new location creates a copy of the script at that location. The original version is still on your OneDrive or other SharePoint location.

> [!IMPORTANT]
> Scripts with [external calls](../develop/external-calls.md) can't be run from SharePoint. You'll receive an error saying "Network access calls are not supported at this time for scripts saved to a SharePoint site".

> [!IMPORTANT]
> Power Automate does **not** support scripts stored on SharePoint at this time.

## Restore deleted scripts

When you delete a script in Excel, it goes to your OneDrive or SharePoint recycle bin. To restore a deleted script, follow the steps listed in [How to recover missing, deleted or corrupted items in SharePoint and OneDrive for work or school](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87). Restoring an **.osts** file returns it to the **All scripts** list.

A deleted script is unshared with the workbook. When you restore a script, it does **not** retain its script access. You will need to share the script again.

Restored scripts still work as expected with Power Automate flows. You don't need to recreate the flow connector.

## File ownership and retention

Office Scripts follow the retention and deletion policies specified by Microsoft OneDrive and Microsoft SharePoint. To learn how to handle scripts that were created and shared by a user being removed from your organization, see [Learn about retention for SharePoint and OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true).

During editing, files are temporarily stored in the browser. You must save the script before closing the Excel window to save it to the OneDrive location. Don't forget to save the file after edits, or else those edits will only be in the browser's version of the file.

## Audit Office Scripts usage at the admin level

Discover who is using Office Scripts in your organization with the compliance center audit log. Details about the audit log are found in [Search the audit log in the Security & Compliance Center](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

To specifically audit Office Scripts related activity as an admin, take the following steps.

1. In a InPrivate browser window (or Incognito or other browser-specific limited-tracking mode), open and log into the [Compliance center](https://compliance.microsoft.com/).
1. Go to the **Audit** page.
1. *(One time only)* On the **Search** tab, select **Start recording user and admin activity**.

    > [!IMPORTANT]
    > It may take an hour or two after turning on recording before all activities across the tenant are recorded.

1. Set the desired search options and press **Search**. Filter **Activities** to **Ran script on workbook** to see any time a script was run. You can also filter the **File, folder, or site** field to `.osts`. This reveals who in your organization is creating or modifying scripts.

    :::image type="content" source="../images/audit-log-example.png" alt-text="A few rows of audit log search results, including the 'Ran script on workbook' action and the upload and modification of an .osts file.":::

## See also

- [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Office Scripts settings in M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Undo the effects of Office Scripts](../testing/undo.md)
