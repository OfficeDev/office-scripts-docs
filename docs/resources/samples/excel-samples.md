---
title: First scripts for Excel
description: A collection of basic code samples to use with Office Scripts in Excel.
ms.date: 02/13/2023
ms.localizationpriority: medium
---

# First scripts for Excel

The following samples are simple scripts for you to try on your own workbooks. They form the building blocks to larger solutions. To use them in Excel:

1. Open a workbook in Excel.
1. Open the **Automate** tab.
1. Select **New Script**.
1. Replace the entire script with the sample of your choice.
1. Select **Run** in the Code Editor's task pane.

## Script basics

These samples demonstrate fundamental building blocks for Office Scripts. Expand these scripts to extend your solution and solve common problems.

### Read and log one cell

This sample reads the value of **A1** and prints it to the console.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### Read the active cell

This script logs the value of the current active cell. If multiple cells are selected, the top-leftmost cell will be logged.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### Add data to a range

This script adds a set of values to a new worksheet. The values start in cell **A1**. The data used in this script is predefined, but could be sourced from other places in or out of the workbook.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // The getData call could be replaced by input from Power Automate or a fetch call.
  const data = getData();

  // Create a new worksheet and switch to it.
  const newWorksheet = workbook.addWorksheet("DataSheet");
  newWorksheet.activate();

  // Get a range matching the size of the data.
  const dataRange = newWorksheet.getRangeByIndexes(
    0,
    0,
    data.length,
    data[0].length);

  // Set the data as the values in the range.
  dataRange.setValues(data);
}

function getData(): string[][] {
  return [["Abbreviation", "State/Province", "Country"],
          ["AL", "Alabama", "USA"],
          ["AK", "Alaska", "USA"],
          ["AZ", "Arizona", "USA"],
          ["AR", "Arkansas", "USA"],
          ["CA", "California", "USA"],
          ["CO", "Colorado", "USA"],
          ["CT", "Connecticut", "USA"],
          ["DE", "Delaware", "USA"],
          ["DC", "District of Columbia", "USA"],
          ["FL", "Florida", "USA"],
          ["GA", "Georgia", "USA"],
          ["HI", "Hawaii", "USA"],
          ["ID", "Idaho", "USA"],
          ["IL", "Illinois", "USA"],
          ["IN", "Indiana", "USA"],
          ["IA", "Iowa", "USA"],
          ["KS", "Kansas", "USA"],
          ["KY", "Kentucky", "USA"],
          ["LA", "Louisiana", "USA"],
          ["ME", "Maine", "USA"],
          ["MD", "Maryland", "USA"],
          ["MA", "Massachusetts", "USA"],
          ["MI", "Michigan", "USA"],
          ["MN", "Minnesota", "USA"],
          ["MS", "Mississippi", "USA"],
          ["MO", "Missouri", "USA"],
          ["MT", "Montana", "USA"],
          ["NE", "Nebraska", "USA"],
          ["NV", "Nevada", "USA"],
          ["NH", "New Hampshire", "USA"],
          ["NJ", "New Jersey", "USA"],
          ["NM", "New Mexico", "USA"],
          ["NY", "New York", "USA"],
          ["NC", "North Carolina", "USA"],
          ["ND", "North Dakota", "USA"],
          ["OH", "Ohio", "USA"],
          ["OK", "Oklahoma", "USA"],
          ["OR", "Oregon", "USA"],
          ["PA", "Pennsylvania", "USA"],
          ["RI", "Rhode Island", "USA"],
          ["SC", "South Carolina", "USA"],
          ["SD", "South Dakota", "USA"],
          ["TN", "Tennessee", "USA"],
          ["TX", "Texas", "USA"],
          ["UT", "Utah", "USA"],
          ["VT", "Vermont", "USA"],
          ["VA", "Virginia", "USA"],
          ["WA", "Washington", "USA"],
          ["WV", "West Virginia", "USA"],
          ["WI", "Wisconsin", "USA"],
          ["WY", "Wyoming", "USA"],
          ["AB", "Alberta", "CAN"],
          ["BC", "British Columbia", "CAN"],
          ["MB", "Manitoba", "CAN"],
          ["NB", "New Brunswick", "CAN"],
          ["NL", "Newfoundland and Labrador", "CAN"],
          ["NT", "Northwest Territory", "CAN"],
          ["NS", "Nova Scotia", "CAN"],
          ["NU", "Nunavut Territory", "CAN"],
          ["ON", "Ontario", "CAN"],
          ["PE", "Prince Edward Island", "CAN"],
          ["QC", "Quebec", "CAN"],
          ["SK", "Saskatchewan", "CAN"],
          ["YT", "Yukon Territory", "CAN"]];
}
```

### Change an adjacent cell

This script gets adjacent cells using relative references. Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### Change all adjacent cells

This script copies the formatting in the active cell to the neighboring cells. Note that this script only works when the active cell isn't on an edge of the worksheet.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### Change each individual cell in a range

This script loops over the currently select range. It clears the current formatting and sets the fill color in each cell to a random color.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### Get groups of cells based on special criteria

This script gets all the blank cells in the current worksheet's used range. It then highlights all those cells with a yellow background.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## Collections

These samples work with collections of objects in the workbook.

### Iterate over collections

This script gets and logs the names of all the worksheets in the workbook. It also sets the their tab colors to a random color.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### Query and delete from a collection

This script creates a new worksheet. It checks for an existing copy of the worksheet and deletes it before making a new sheet.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## Display data

These samples demonstrate how to work with worksheet data and provide users with a better view or organization.

### Apply conditional formatting

This sample applies conditional formatting to the currently used range in the worksheet. The conditional formatting is a green fill for the top 10% of values.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### Log the "Grand Total" values from a PivotTable

This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="A PivotTable showing fruit sales with the Grand Total row highlighted green.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## Formulas

These samples use Excel formulas and show how to work with them in scripts.

### Single formula

This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### Handle a `#SPILL!` error returned from a formula

This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function. If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

### Replace all formulas with their result values

This script replaces every cell in the current worksheet that contains a formula with the result of that formula. This means there won't be any formulas after the script is run, only values.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the ranges with formulas.
    let sheet = workbook.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaCells = usedRange.getSpecialCells(ExcelScript.SpecialCellType.formulas);

    // In each formula range: get the current value, clear the contents, and set the value as the old one.
    // This removes the formula but keeps the result.
    formulaCells.getAreas().forEach((range) => {
      let currentValues = range.getValues();
      range.clear(ExcelScript.ClearApplyTo.contents);
      range.setValues(currentValues);
    });
}
```

## Suggest new samples

We welcome suggestions for new samples. If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.

## See also

* [Sudhi Ramamurthy's "Range basics" on YouTube](https://youtu.be/4emjkOFdLBA)
* [Office Scripts samples and scenarios](samples-overview.md)
* [Record, edit, and create Office Scripts in Excel](../../tutorials/excel-tutorial.md)
