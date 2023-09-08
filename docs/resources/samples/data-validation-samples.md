---
title: Create a drop-down list using data validation
description: Learn how to add data validation to a cell and give the user a selection of values to enter.
ms.date: 09/08/2023
ms.localizationpriority: medium
---

# Create a drop-down list using data validation

This script creates a drop-down selection list for a cell. It uses the existing values of the selected range as the choices for the list.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="A worksheet showing a range of three cells containing color choices 'red, blue, green' and next to it, the same choices shown in a drop-down list.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

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
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```
