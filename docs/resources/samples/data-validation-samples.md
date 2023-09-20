---
title: "Data validation: dropdown lists, prompts, and warning pop-ups"
description: Learn how to add data validation to a cell and give the user a selection of values to enter.
ms.date: 09/20/2023
ms.localizationpriority: medium
---

# Data validation: dropdown lists, prompts, and warning pop-ups

Data validation helps the user ensure consistency in a worksheet. Use these features to limit what can be entered into a cell and provide warnings or errors to users when those conditions aren't met. To learn more about data validation in Excel, see [Apply data validation to cells](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249).

## Create a dropdown list using data validation

The following sample creates a dropdown selection list for a cell. It uses the existing values of the selected range as the choices for the list.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="A worksheet showing a range of three cells containing color choices 'red, blue, green' and next to it, the same choices shown in a dropdown list.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  const selectedRange = workbook.getSelectedRange();
  const rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  const targetCell = selectedRange.getCell(0,0);
  const dataValidation = targetCell.getDataValidation();

  // Set the content of the dropdown list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## Add a prompt to a range

This example creates a prompt note that appears when a user enters the given cells. This is used to remind users about input requirements, without strict enforcement.

:::image type="content" source="../../images/data-validation-prompt.png" alt-text="A prompt with the title 'First names only' and the message 'Only enter the first name of the employee, not the full name.' next to a worksheet with some names in cells.":::

```TypeScript
/**
 * This script creates a text prompt that's shown in C2:C8 when a user enters the cell.
 */
function main(workbook: ExcelScript.Workbook) {
    // Get the data validation object for C2:C8 in the current worksheet.
    const selectedSheet = workbook.getActiveWorksheet();
    const dataValidation = selectedSheet.getRange("C2:C8").getDataValidation();

    // Clear any previous validation to avoid conflicts.
    dataValidation.clear();

    // Create a prompt to remind users to only enter first names in this column.
    const prompt: ExcelScript.DataValidationPrompt = {
      showPrompt: true,
      title: "First names only",
      message: "Only enter the first name of the employee, not the full name."
    }
    dataValidation.setPrompt(prompt);
}
```

## Alert the user when invalid data is entered

The following sample script prevents the user from entering anything other than positive numbers into a range. If they try to put anything else, an error message pops up and indicates the problem.

:::image type="content" source="../../images/data-validation-error.png" alt-text="An error message with the title 'Invalid data' and the message 'Positive numbers only.' next to a cell with a negative number.":::

```TypeScript
/**
 * This script creates a data validation rule for the range B2:B5.
 * All values in that range must be a positive number.
 * Attempts to enter other values are blocked and an error message appears.
 */
function main(workbook: ExcelScript.Workbook) {
    // Get the range B2:B5 in the active worksheet.
    const currentSheet = workbook.getActiveWorksheet();
    const positiveNumberOnlyCells = currentSheet.getRange("B2:B5");

    // Create a data validation rule to only allow positive numbers.
    const positiveNumberValidation: ExcelScript.BasicDataValidation = {
        formula1: "0",
        operator: ExcelScript.DataValidationOperator.greaterThan
    };
    const positiveNumberOnlyRule: ExcelScript.DataValidationRule = {
      wholeNumber: positiveNumberValidation
    };

    // Set the rule on the range.
    const rangeDataValidation = positiveNumberOnlyCells.getDataValidation();
    rangeDataValidation.setRule(positiveNumberOnlyRule);

    // Create an alert to appear when data other than positive numbers are entered.
    const positiveNumberOnlyAlert: ExcelScript.DataValidationErrorAlert = {
        message: "Positive numbers only.",
        showAlert: true,
        style: ExcelScript.DataValidationAlertStyle.stop,
        title: "Invalid data"
    };
    rangeDataValidation.setErrorAlert(positiveNumberOnlyAlert);
}
```
