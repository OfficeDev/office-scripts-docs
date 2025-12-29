---
title: Set conditional formatting for cross-column comparisons
description: Learn how to apply conditional formatting and get input from the user.
ms.date: 12/22/2025
ms.localizationpriority: medium
---

# Set conditional formatting for cross-column comparisons

This sample shows how to apply conditional formatting to a range. The conditions used are comparing values to those in an adjacent column.  Additionally, this sample uses [parameters to get user input](../../develop/user-input.md). This lets the person running the script select the range, the type of comparison, and the colors.

## Sample code: Set conditional formatting

[!INCLUDE [open-code-editor-single-script](../../includes/open-code-editor-single-script.md)]

```TypeScript
/**
 * Formats a range on the current sheet based on values in an adjacent column.
 * @param rangeAddress The A1-notation range to format.
 * @param compareTo The adjacent column to compare against.
 * @param colorIfGreater The color of the cell if the value is greater than the adjacent column.
 * @param colorIfEqual The color of the cell if the value is equal to the adjacent column.
 * @param colorIfLess The color of the cell if the value is less than the adjacent column.
 */
function main(
  workbook: ExcelScript.Workbook,
  rangeAddress: string, compareTo: "Left" | "Right",
  colorIfGreater: "Red" | "Green" | "Yellow" | "None",
  colorIfLess: "Red" | "Green" | "Yellow" | "None",
  colorIfEqual: "Red" | "Green" | "Yellow" | "None"
) {
  // Get the specified range.
  const selectedSheet = workbook.getActiveWorksheet();
  const range = selectedSheet.getRange(rangeAddress);

  // Remove old conditional formatting.
  range.clearAllConditionalFormats();

  // Get the address of the first adjacent cell of the adjacent column.
  let adjacentColumn: string;
  if (compareTo == "Left") {
    adjacentColumn = range.getColumnsBefore().getCell(0, 0).getAddress();
  } else {
    adjacentColumn = range.getColumnsAfter().getCell(0, 0).getAddress();
  }

  // Remove the worksheet name from the address to create a relative formula.
  let formula = "=$" + adjacentColumn.substring(adjacentColumn.lastIndexOf("!") + 1);

  // Set the conditional formatting based on the user's color choices.
  setConditionalFormatting(
    range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue).getCellValue(),
    colorIfGreater, 
    formula, 
    ExcelScript.ConditionalCellValueOperator.greaterThan);
  setConditionalFormatting(
    range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue).getCellValue(),
    colorIfEqual, 
    formula, 
    ExcelScript.ConditionalCellValueOperator.equalTo);
  setConditionalFormatting(
    range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue).getCellValue(),
    colorIfLess, 
    formula, 
    ExcelScript.ConditionalCellValueOperator.lessThan);
}

function setConditionalFormatting(
  conditionalFormat: ExcelScript.CellValueConditionalFormat,
   color: "Red" | "Green" | "Yellow" | "None", 
   formula: string, 
   operator: ExcelScript.ConditionalCellValueOperator
) {
  // Pick the fill and font colors based on the preset color choices.
  if (color == "Red") {
    conditionalFormat.getFormat().getFont().setColor("#9C0006");
    conditionalFormat.getFormat().getFill().setColor("#FFC7CE");
  } else if (color == "Green") {
    conditionalFormat.getFormat().getFont().setColor("#001600");
    conditionalFormat.getFormat().getFill().setColor("#C6EFCE");
  } else if (color == "Yellow") {
    conditionalFormat.getFormat().getFont().setColor("#9C5700");
    conditionalFormat.getFormat().getFill().setColor("#FFEB9C");
  } else { /* None */
    return;
  }

  // Apply the conditional formatting.
  conditionalFormat.setRule({ 
    formula1: formula,
    operator: operator
  });
}
```
