---
title: Run Office Scripts in Excel with buttons
description: Add buttons to workbooks that control Office Scripts in Excel.
ms.topic: overview
ms.date: 03/10/2023
ms.localizationpriority: medium
---

# Run Office Scripts in Excel with buttons

Help your colleagues find and run your scripts by adding script buttons to a workbook.

:::image type="content" source="../images/run-from-button.png" alt-text="A button in the worksheet that runs a script when clicked.":::

## Create script buttons

With any script, go to the **More options (…)** menu in either the script's details page or the Code Editor's task pane and select **Add button**. This creates a button in the workbook that runs the associated script when selected. It also shares the script with the workbook, so everyone with write permissions to the workbook can use your helpful automation.

The following screenshot shows the script details page for a script titled **Create PivotTable** and has the **Add button** option within the **More options (…)** menu highlighted.

:::image type="content" source="../images/add-button.png" alt-text="The 'Add button' option in the script details page menu.":::

## Remove script buttons

To stop sharing a script through a button, go to the **More options (…)** menu in the script details page and select **Stop sharing**. This removes all the buttons that run the script. Deleting a single button removes the script from that one button, even if the operation is undone or the button is cut and pasted.

## Older versions of Excel

Script buttons won't work when opened in [versions of Excel that don't support Office Scripts](../testing/platform-limits.md#platform-support). In that case, the button still appears, but selecting it has no effect.

## See also

- [Platform limits and requirements with Office Scripts](../testing/platform-limits.md)
