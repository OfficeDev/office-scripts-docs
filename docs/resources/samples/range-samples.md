---
title: 'Ranges: Work with the grid in Office Scripts'
description: A collection of code samples showing how to work with Excel ranges, collections of cells, in Office Scripts.
ms.date: 12/05/2023
ms.localizationpriority: medium
---

# Ranges: Work with the grid in Office Scripts

The following samples are simple scripts for you to try on your own workbooks. They form the building blocks to larger solutions. Expand these scripts to extend your solution and solve common problems.

## Read and log one cell

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

## Read the active cell

This script logs the value of the current active cell. If multiple cells are selected, the top-leftmost cell will be logged.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

## Add data to a range

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

## Change an adjacent cell

This script gets adjacent cells using relative references. Note that if the active cell is on the top row, part of the script fails because it references the cell above the currently selected one.

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

## Change all adjacent cells

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

## Change each individual cell in a range

This script loops over the currently selected range. It clears the current formatting and sets the fill color in each cell to a random color.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting.
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random hexadecimal color code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hexadecimal code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

## Get groups of cells based on special criteria

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

## Formulas

Ranges have both values and formulas. The formula is the expression to be evaluated. The value is the result of that expression.

### Single formula

This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1");
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

We welcome suggestions for new samples. If there's a common scenario that would help other script developers, please tell us in the Feedback section at the bottom of the page.

## See also

* [Sudhi Ramamurthy's "Range basics" on YouTube](https://youtu.be/4emjkOFdLBA)
* [Chart samples](chart-samples.md)
* [Office Scripts samples and scenarios](samples-overview.md)
* [Tutorial: Create and format an Excel table](../../tutorials/excel-tutorial.md)
