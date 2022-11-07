---
title: Office Scripts in Excel
description: A brief introduction to the Action Recorder and Code Editor for Office Scripts.
ms.topic: overview
ms.date: 10/05/2022
ms.localizationpriority: high
---

# Office Scripts in Excel

Office Scripts in Excel let you automate your day-to-day tasks. You can create and edit scripts with the Code Editor. Run a series of Excel steps with a single button press. Then, share that script with coworkers so everyone can improve their workflow. Inside Excel on the web, you have an additional tool: the Action Recorder. This turns your manual steps into Office Scripts without having to look at any code.

This series of documents teaches you how to use these tools. You'll find a wealth of samples covering different Excel scenarios. Use the tutorials to introduce yourself to the Action Recorder and Code Editor. These provide step-by-step guidance on how to record your frequent Excel actions, edit those scripts, and create new scripts from scratch.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## Requirements

To use Office Scripts, you'll need the following.

1. Excel for Windows or Excel on the web.

    > [!TIP]
    > Office Scripts are now available in Office on Mac for [Office Insiders](https://insider.office.com/).

1. OneDrive for Business.
1. Any commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps, such as:

    - Office 365 Business
    - Office 365 Business Premium
    - Office 365 ProPlus
    - Office 365 ProPlus for Devices
    - Office 365 Enterprise E3
    - Office 365 Enterprise E5
    - Office 365 A3
    - Office 365 A5

> [!NOTE]
> If you meet these requirements and are still not seeing the **Automate** tab, it's possible that your admin has disabled the feature or there's some other problem with your environment. Please follow the steps under [Automate tab not appearing or Office Scripts unavailable](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable) to start using Office Scripts.

## When to use Office Scripts

Scripts allow you to record and replay your Excel actions on different workbooks and worksheets. If you find yourself doing the same things over and over again, you can turn all that work into an easy-to-run Office Script. Run your script with a button-press in Excel or combine it with Power Automate to streamline your entire workflow.

As an example, say you start your work day by opening a .csv file from an accounting site in Excel. You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet. Those actions you repeat daily can be recorded once with the Action Recorder. From then on, running the script will take care of your entire .csv conversion. You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything. Office Scripts allows you to automate your common tasks so you and your workplace can be more efficient and productive.

## Action Recorder (web-only)

:::image type="content" source="../images/action-recorder-intro.png" alt-text="A list of actions recorded by Action Recorder.":::

The Action Recorder records actions you take in Excel and saves them as a script. With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables. The resulting script can be run on other worksheets and workbooks to recreate your original actions.

## Code Editor

:::image type="content" source="../images/code-editor-intro.png" alt-text="The Code Editor showing the script code used in this tutorial.":::

Use the Code Editor to edit scripts recorded with the Action Recorder or make a brand to script. This tool lets you tweak and customize scripts to better suit your exact needs. You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.

> [!TIP]
> The Action Recorder has a **Copy as code** button to record actions into script code without saving the entire script.
>
> :::image type="content" source="../images/action-recorder-copy-code.png" alt-text="The Action Recorder task pane with the 'Copy as code' button highlighted.":::

Our [tutorials](../tutorials/excel-tutorial.md) provide a guided and structured way learn the capabilities of Office Scripts. After completing the tutorials, read [Scripting fundamentals for Office Scripts in Excel](../develop/scripting-fundamentals.md) to learn more about the Code Editor and how to write and edit your own scripts. For additional information about the Code Editor and how your script code is interpreted, read [Office Scripts Code Editor environment](code-editor-environment.md).

## Share Office Scripts

Office Scripts can be shared with other users of an Excel workbook. When you share a script in a shared workbook, everyone with access to the workbook can also view and run your script. For more details about sharing and unsharing scripts, see [Sharing Office Scripts in Excel](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b).

:::image type="content" source="../images/script-sharing.png" alt-text="The script details page showing the 'Share with others in this workbook' option.":::

Add buttons that run scripts to help your colleagues discover your valuable solutions and let them run scripts straight from the workbook. Learn more about script buttons in [Run Office Scripts with buttons](../develop/script-buttons.md).

:::image type="content" source="../images/add-button.png" alt-text="A button in the worksheet that runs a script when clicked.":::

> [!NOTE]
> Learn more about how scripts are stored in your OneDrive in [Office Scripts file storage and ownership](script-storage.md).

## Connect Office Scripts to Power Automate

[Power Automate](https://flow.microsoft.com/) is a service that helps you create automated workflows between multiple apps and services. Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook. You can run your scripts on a schedule, trigger them in response to emails, and much more. Visit the [Run Office Scripts with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.

## Next steps

Complete the [Office Scripts in Excel tutorial](../tutorials/excel-tutorial.md) to learn how to create your first script.

## See also

- [Scripting fundamentals for Office Scripts in Excel](../develop/scripting-fundamentals.md)
- [Office Scripts API reference](/javascript/api/office-scripts/overview)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Office Scripts settings in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Introduction to Office Scripts in Excel](https://support.microsoft.com/office/9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Sharing Office Scripts in Excel](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office Scripts Dev Center](https://developer.microsoft.com/office-scripts)
