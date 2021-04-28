---
title: 'Manage calculation mode in Excel'
description: 'Learn how to use Office Scripts to manage the calculation mode in Excel on the web.'
ms.date: 04/28/2021
localization_priority: Normal
---

# Manage calculation mode in Excel

This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts. You can try the script on any Excel file.

## Scenario

In Excel on the web, a file's calculation mode can be controlled programmatically using APIs. The following actions are possible using Office Scripts.

1. Get the calculation mode.
1. Set the calculation mode.
1. Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).

## Sample code: Control calculation mode

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## Training video: Manage calculation mode

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).
