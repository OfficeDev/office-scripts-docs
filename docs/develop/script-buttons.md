---
title: 'Run Office Scripts with buttons'
description: Add buttons to workbooks that control Office Scripts.
ms.topic: overview
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# Run Office Scripts with buttons

Help your colleagues find and run your scripts by adding script buttons to a workbook.

:::image type="content" source="../images/run-from-button.png" alt-text="A button in the worksheet that runs a script when clicked.":::

## Create script buttons

With any script, go to the **More options (…)** menu in either the script's Details page or the Code Editor's task pane and select **Add button**. This creates a button in the workbook that runs the associated script when selected. It also shares the script with the workbook, so everyone with write permissions to the workbook can use your helpful automation.

The following screenshot shows the script Details page for a script titled **Create PivotTable** and has the **Add button** option within the **More options (…)** menu highlighted.

:::image type="content" source="../images/add-button.png" alt-text="The 'Add button' option in the script Details page menu.":::

## Remove script buttons

To stop sharing a script through a button, go to the **More options (…)** menu in the script's Details page and select **Stop sharing**. This removes all the buttons that run the script. Deleting a single button removes the script from that one button, even if the operation is undone or the button is cut and pasted.

## Script buttons on Excel for Windows

These script buttons also work on Windows. Create the button in Excel on the web and users on Windows can run your script with the click of a button. Please note that only running the button is supported on Windows. If you need to edit the script, you must do that through the web application.

> [!NOTE]
> This feature is being rolled out to users with a Microsoft 365 subscription and is not available to everyone. It's slowly released to larger numbers of users to ensure that it's working as expected. This feature is subject to change based on your feedback. Unsupported platforms or Office versions without the feature will display the shape used for the script button, but the button won't be clickable.
