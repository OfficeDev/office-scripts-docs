---
title: Basic scripts for Office Scripts in Excel
description: A collection of code samples to use with Office Scripts in Excel.
ms.date: 02/13/2023
ms.localizationpriority: medium
---

# Basic scripts for Office Scripts in Excel

The following samples are simple scripts for you to try on your own workbooks. To use them in Excel:

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

This script adds a set of values to a new worksheet. The values start in cell **A1**. The data used in this script is pre-defined, but could be sourced from other places in or out of the workbook.

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

## Row and column visibility

These samples demonstrate how to show, hide, and freeze rows and columns.

### Hide columns

This script hides columns "D", "F", and "J".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  // Hide columns D, F, and J.
  sheet.getRange("D:D").setColumnHidden(true);
  sheet.getRange("F:F").setColumnHidden(true);
  sheet.getRange("J:J").setColumnHidden(true);
}
```

### Show all rows and columns

This script get the worksheet's used range, checks if there are any hidden rows and columns, and shows them.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the currently selected sheet.
    const selectedSheet = workbook.getActiveWorksheet();

    // Get the entire data range.
    const range = selectedSheet.getUsedRange();

    // If the used range is empty, end the script.
    if (!range) {
      console.log(`No data on this sheet.`)
      return;
    }

    // If no columns are hidden, log message, else, show columns
    if (range.getColumnHidden() == false) {
      console.log(`No columns hidden`);
    } else {
      range.setColumnHidden(false);
    }

    // If no rows are hidden, log message, else, show rows.
    if (range.getRowHidden() == false) {
      console.log(`No rows hidden`);
    } else {
      range.setRowHidden(false);
    }
}
```

### Freeze currently selected cells

This script checks what cells are currently selected and freezes that selection, so those cells are always visible.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the currently selected sheet.
    const selectedSheet = workbook.getActiveWorksheet();

    // Get the current selected range.
    const selectedRange = workbook.getSelectedRange();

    // If no cells are selected, end the script. 
    if (!selectedRange) {
      console.log(`No cells in the worksheet are selected.`);
      return;
    }

    // Log the address of the selected range
    console.log(`Selected range for the worksheet: ${selectedRange.getAddress()}`);

    // Freeze the selected range.
    selectedSheet.getFreezePanes().freezeAt(selectedRange);
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

## Dates

The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.

The following sample gets the current date and time and then writes those values to two cells in the active worksheet.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object. It uses the date's numeric serial number as input for the JavaScript Date. This serial number is described in the [NOW() function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) article.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## Tables

The samples in this section showcase common interactions with Excel tables.

### Create a sorted table

This sample creates a table from the current worksheet's used range, then sorts it based on the first column.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### Filter a table

This sample filters an existing table using the values in one of the columns.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table in the workbook named "StationTable".
  const table = workbook.getTable("StationTable");

  // Get the "Station" table column for the filter.
  const stationColumn = table.getColumnByName("Station");

  // Apply a filter to the table that will only show rows 
  // with a value of "Station-1" in the "Station" column.
  stationColumn.getFilter().applyValuesFilter(["Station-1"]);
}
```

> [!TIP]
> Copy the filtered information across the workbook by using `Range.copyFrom`. Add the following line to the end of the script to create a new worksheet with the filtered data.
>
> ```TypeScript
>   workbook.addWorksheet().getRange("A1").copyFrom(table.getRange());
> ```

### Dynamically reference table values

This script uses the "@COLUMN_NAME" syntax to set formulas in a table column. The column names in the table can be changed without changing this script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const table = workbook.getTable("Profits");

  // Get the column names for columns 2 and 3.
  // Note that these are 1-based indices.
  const nameOfColumn2 = table.getColumn(2).getName();
  const nameOfColumn3 = table.getColumn(3).getName();

  // Set the formula of the fourth column to be that row's column 2 and 3 values.
  const combinedColumn = table.getColumn(4).getRangeBetweenHeaderAndTotal();
  combinedColumn.setFormula(`=[@[${nameOfColumn2}]]*[@[${nameOfColumn3}]]`)
}
```

#### Before the script

| Month | Price | Units Sold | Total |
|--|--|--|--|
| Jan | 45 | 5 |  |
| Feb | 45 | 3 |  |
| Mar | 45 | 6 |  |

#### After the script

| Month | Price | Units Sold | Total |
|--|--|--|--|
| Jan | 45 | 5 | 225 |
| Feb | 45 | 3 | 135 |
| Mar | 45 | 6 | 270 |

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

### Create a drop-down list using data validation

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
