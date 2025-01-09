---
title: Office Scripts in Excel
description: A brief introduction to the Action Recorder and Code Editor for Office Scripts.
ms.topic: overview
ms.date: 08/08/2024
ms.localizationpriority: high
---

# Office Scripts in Excel

Office Scripts in Excel let you automate your day-to-day tasks. Use the Action Recorder to turn manual steps into reusable scripts. Edit those scripts or create new ones with the Code Editor. Let others in the workbook run these scripts with a single button. Then, share them with coworkers so everyone can improve their workflow.

This series of documents teaches you how to use these tools. You'll find a wealth of samples covering different Excel scenarios. Use the tutorials to introduce yourself to the Action Recorder and Code Editor. These provide step-by-step guidance on how to record your frequent Excel actions, edit those scripts, and create new scripts from scratch.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## When to use Office Scripts

Scripts allow you to record and replay your Excel actions on different workbooks and worksheets. If you find yourself doing the same things over and over again, you can turn all that work into an easy-to-run Office Script. Run your script with a button in Excel or combine it with Power Automate to streamline your entire workflow.

As an example, imagine at the start of each work day you open a .csv file from an accounting site in Excel. You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet. Those actions you repeat daily can be recorded once with the Action Recorder. From then on, running the script will take care of your entire .csv conversion. You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything. Office Scripts allows you to automate your common tasks so you and your workplace can be more efficient and productive.

## Action Recorder

:::image type="content" source="../images/action-recorder-intro.png" alt-text="A list of actions recorded by Action Recorder.":::

The Action Recorder records actions you take in Excel and saves them as a script. With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables. The resulting script can be run on other worksheets and workbooks to recreate your original actions.

More information about the Action Recorder can be found in the article [Record your actions as Office Script](https://support.microsoft.com/office/453ab58d-708f-40a9-ab87-99a743bfa69a).

## Code Editor

:::image type="content" source="../images/code-editor-intro.png" alt-text="The Code Editor showing the script code used in this tutorial.":::

Use the Code Editor to edit scripts recorded with the Action Recorder or make a brand new script. This tool lets you tweak and customize scripts to better suit your exact needs. You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.

> [!TIP]
> The Action Recorder has a **Copy as code** button to record actions into script code without saving the entire script.
>
> :::image type="content" source="../images/action-recorder-copy-code.png" alt-text="The Action Recorder task pane with the 'Copy as code' button highlighted.":::

Our [tutorials](../tutorials/excel-tutorial.md) provide a guided and structured way learn the capabilities of Office Scripts. After completing the tutorials, read [Fundamentals for Office Scripts in Excel](../develop/scripting-fundamentals.md) to learn more about the Code Editor and how to write and edit your own scripts. For additional information about the Code Editor and how your script code is interpreted, read [Office Scripts Code Editor environment](code-editor-environment.md).

## Share Office Scripts

Office Scripts can be shared with other users in your organization. When you share a script in a shared workbook, team members with access to the workbook can also view and run your script. For more details about sharing and unsharing scripts, see [Sharing Office Scripts in Excel](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b).

Add buttons that run scripts to help your colleagues discover your valuable solutions and let them run scripts straight from the workbook. Learn more about script buttons in [Run Office Scripts with buttons](../develop/script-buttons.md).

:::image type="content" source="../images/add-button.png" alt-text="The 'Add in workbook' button on the 'Create Report' script details page with a button named 'Create Report' shown in the Excel grid.":::

> [!NOTE]
> Learn more about how scripts are stored in your OneDrive in [Office Scripts file storage and ownership](script-storage.md).

## Schedule scripts to run automatically

> [!IMPORTANT]
> Script scheduling is temporarily disabled within Office Scripts. Existing scheduled scripts will continue to run. In the interim, use Power Automate to create a flow and schedule your scripts to run in that flow. To learn more, see [Run scripts with Power Automate](../develop/power-automate-integration.md).

Set your scripts to run every day and keep your workbook up-to-date. Once you have your script, you can set it to automatically run on the workbook at regular intervals. A behind-the-scenes Power Automate flow ensures everything happens, even when the workbook is closed.

To schedule a script, open the script in the Code Editor. Open the **Script scheduling** section and complete the sign in process to Excel through Power Automate. Set how often you want the script to run and select **Create flow** to begin.

:::image type="content" source="../images/schedule-a-script.png" alt-text="The Code Editor task pane that shows the recurrence interval options for scheduling a script.":::

## Connect Office Scripts to Power Automate

[Power Automate](https://make.powerautomate.com/) is a service that helps you create automated workflows between multiple apps and services. Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook. You can run your scripts on a schedule, trigger them in response to emails, and much more. Visit the [Run Office Scripts with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.

## Next steps

Complete the [Office Scripts in Excel tutorial](../tutorials/excel-tutorial.md) to learn how to create your first script.

## See also

- [Fundamentals for Office Scripts in Excel](../develop/scripting-fundamentals.md)
- [Office Scripts samples and scenarios](../resources/samples/samples-overview.md)
- [Office Scripts API reference](/javascript/api/office-scripts/overview)
- [Platform limits and requirements with Office Scripts](../testing/platform-limits.md)
- [Office Scripts settings in M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Sharing Office Scripts in Excel](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
