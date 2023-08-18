---
title: Pass data to and from scripts in Power Automate
description: How and why to use parameters and return values in Office Scripts with Power Automate.
ms.date: 08/15/2023
ms.localizationpriority: medium
---

# Pass data to and from scripts in Power Automate

Power Automate chains together separate programs into a single automated workflow. Each connector has different parameters it accepts and different values it returns. Your scripts can be written to expand the "Run script" Power Automate action to get additional input or give output.

Input for your script is specified by adding parameters to the `main` function. Output from the script is declared by adding a return type to `main`.

> [!NOTE]
> When you create a "Run script" block in your flow, the accepted parameters and returned types are populated. If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow. This ensures the data is being parsed correctly.

## Pass data to scripts with parameters

Add parameters to scripts to provide input from other parts of the flow. It's the same process to add parameters for flow-based scripts as it is for scripts run through the Excel client. Learn about providing input to scripts in [Get user input for scripts](user-input.md).

The following screenshot shows what a script with the signature `function main(workbook: ExcelScript.Workbook, location: string = "Seattle")` would display.

:::image type="content" source="../images/power-automate-default-parameter.png" alt-text="The Run script action showing an additional parameter field called 'Location' with a pre-populated value of 'Seattle'.":::

The [dropdown menus created by type unions](user-input.md#dropdown-lists-for-parameters) also function the same in Power Automate.

:::image type="content" source="../images/power-automate-drop-down-parameter-choices.png" alt-text="The Run script action showing an additional parameter field called 'Location' with choices between 'Seattle' and 'Redmond'.":::

## Return data from a script

Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow. To return an object, add the return type syntax to the `main` function. For example, if you wanted to return a `string` value from the script, your `main` signature would be `function main(workbook: ExcelScript.Workbook): string`.

Returned values are shown as dynamic content from the Run script action in the flow. The dynamic content is named "result".

:::image type="content" source="../images/power-automate-return-dynamic-content.png" alt-text="The dynamic content selector in Power Automate showing an entry from a Run script action named 'result'.":::

Acceptable types for returning data are the same as for parameters. Details on type restrictions are found in the article [Get user input for scripts](user-input.md#type-restrictions).

## See also

- [Call scripts from a manual Power Automate flow](../tutorials/excel-power-automate-manual.md)
- [Get user input for scripts](user-input.md)
- [Run Office Scripts with Power Automate](power-automate-integration.md)
- [Troubleshooting information for Power Automate with Office Scripts](../testing/power-automate-troubleshooting.md)
