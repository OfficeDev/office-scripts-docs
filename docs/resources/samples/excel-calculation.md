---
title: Manage calculation mode in Excel
description: Learn how to use Office Scripts to manage the calculation mode in Excel.
ms.date: 12/22/2025
ms.localizationpriority: medium
---

# Manage calculation mode in Excel

This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel using Office Scripts. You can try the script on any Excel file.

## Scenario

Workbooks with large numbers of formulas can take a while to recalculate. Rather than letting Excel control when calculations happen, you can manage them as part of your script. This will help with performance in certain scenarios.

The sample script sets the calculation mode to manual. This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4)). The script then displays the current calculation mode and fully recalculates the entire workbook.

## Sample code: Control calculation mode

[!INCLUDE [open-code-editor-single-script](../../includes/open-code-editor-single-script.md)]

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## Training video: Manage calculation mode

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).
