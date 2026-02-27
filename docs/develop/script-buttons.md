---
title: Run Office Scripts in Excel from buttons
description: Add buttons to workbooks that control Office Scripts in Excel.
ms.topic: overview
ms.date: 02/26/2026
ms.localizationpriority: medium
---

# Run Office Scripts in Excel from buttons

Help your colleagues find and run your scripts by adding script buttons to a workbook.

:::image type="content" source="../images/run-from-button.png" alt-text="A button in the worksheet that runs a script when clicked.":::

> [!NOTE]
> Script buttons aren't available during the Office Scripts preview for personal and family Microsoft 365 subscriptions.

## Create script buttons

When viewing a script, scroll to the **Share this script** section and enable the **Associate with workbook** toggle. This shares the script with the workbook, so everyone with write permissions to the workbook can use your helpful automation. Select **Add button to worksheet** to create a button in the worksheet that runs the associated script.

:::image type="content" source="../images/add-button.png" alt-text="The 'Associate with workbook' and 'Add button to worksheet' controls in the Share this script section.":::

> [!IMPORTANT]
> If OneDrive sharing is restricted by organizational policies, you can't associate a script or create a script button.

## Remove script buttons

To stop sharing a script, disable the **Associate with workbook** toggle in the **Share this script** section of the script details page. This removes all the buttons that run the script from the workbook.

:::image type="content" source="../images/associate-with-workbook.png" alt-text="The 'Share this script' section of the script details page, with the 'Associate with workbook' toggle enabled.":::

> [!IMPORTANT]
> Deleting a script button from the Excel grid removes that button, but the script still remains associated with the workbook.

## Older versions of Excel

Script buttons won't work when opened in [versions of Excel that don't support Office Scripts](../testing/platform-limits.md#platform-support). In that case, the button still appears, but selecting it has no effect.

## See also

- [Platform limits and requirements with Office Scripts](../testing/platform-limits.md)
