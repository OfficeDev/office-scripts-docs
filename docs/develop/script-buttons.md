---
title: Run Office Scripts in Excel from buttons
description: Add buttons to workbooks that control Office Scripts in Excel.
ms.topic: overview
ms.date: 06/06/2024
ms.localizationpriority: medium
---

# Run Office Scripts in Excel from buttons

Help your colleagues find and run your scripts by adding script buttons to a workbook.

:::image type="content" source="../images/run-from-button.png" alt-text="A button in the worksheet that runs a script when clicked.":::

> [!NOTE]
> Script buttons aren't available during the Office Scripts preview for personal and family Microsoft 365 subscriptions.

## Create script buttons

When viewing a script, select **Add in workbook**. This creates a button in the workbook that runs the associated script. It also shares the script with the workbook, so everyone with write permissions to the workbook can use your helpful automation.

:::image type="content" source="../images/add-button.png" alt-text="The 'Add in workbook' button on the 'Create Report' script details page with a button named 'Create Report' shown in the Excel grid.":::

> [!IMPORTANT]
> If OneDrive sharing is restricted by organizational policies, you can't create a script button.

## Remove script buttons

To stop sharing a script through a button, select the arrow next to **Add in workbook** and choose the option **Remove from workbook**. This removes all the buttons that run the script. Deleting a single button removes the script from that one button, even if the operation is undone or the button is cut and pasted.

:::image type="content" source="../images/remove-button.png" alt-text="The 'Remove from workbook' option on the script details page.":::

## Older versions of Excel

Script buttons won't work when opened in [versions of Excel that don't support Office Scripts](../testing/platform-limits.md#platform-support). In that case, the button still appears, but selecting it has no effect.

## See also

- [Platform limits and requirements with Office Scripts](../testing/platform-limits.md)
