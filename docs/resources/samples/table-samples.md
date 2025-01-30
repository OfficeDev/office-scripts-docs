---
title: Table samples
description: A collection of samples showing how to interact with Excel tables.
ms.date: 09/25/2024
ms.localizationpriority: medium
---

# Table samples

These samples showcase common interactions with Excel tables.

## Create a sorted table

This sample creates a table from the current worksheet's used range, then sorts it based on the first column.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  const usedRange = selectedSheet.getUsedRange();
  const newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

## Filter a table

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

### Filter out one value

The previous sample filters a table based on a list of included values. To exclude a particular value from the table, you need to provide the list of every other value in the column. This sample uses a function `columnToSet` to convert a column into a set of unique values. That set then has the excluded value ("Station-1") removed.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table in the workbook named "StationTable".
  const table = workbook.getTable("StationTable");

  // Get the "Station" table column for the filter.
  const stationColumn = table.getColumnByName("Station");

  // Get a list of unique values in the station column.
  const stationSet = columnToSet(stationColumn);

  // Apply a filter to the table that will only show rows
  // that don't have a value of "Station-1" in the "Station" column. 
  stationColumn.getFilter().applyValuesFilter(stationSet.filter((value) => {
      return value !== "Station-1";
  }));
}

/**
 * Convert a column into a set so it only contains unique values.
 */
function columnToSet(column: ExcelScript.TableColumn): string[] {
    const range = column.getRangeBetweenHeaderAndTotal().getValues() as string[][];
    const columnSet: string[] = [];
    range.forEach((value) => {
        if (!columnSet.includes(value[0])) {
            columnSet.push(value[0]);
        }
    });

    return columnSet;
}
```

## Remove table column filters

This sample removes the filters from a table column, based on the active cell location. The script detects if the cell is part of a table, determines the table column, and clears any filters that are applied on it.

Download [table-with-filter.xlsx](table-with-filter.xlsx) for a ready-to-use workbook. Add the following script to try the sample yourself!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there's no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

### Before clearing column filter (notice the active cell)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="An active cell before clearing column filter.":::

### After clearing column filter

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="An active cell after clearing column filter.":::

> [!TIP]
> If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.

## Dynamically reference table values

This script uses the "@COLUMN_NAME" syntax to set formulas in a table column. The column names in the table can be changed without changing this script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const table = workbook.getTable("Profits");

  // Get the column names for columns 2 and 3.
  // Note that these are 1-based indices.
  const nameOfColumn2 = table.getColumn(2).getName();
  const nameOfColumn3 = table.getColumn(3).getName();

  // Set the formula of the fourth column to be the product of the values found
  // in that row's second and third columns.
  const combinedColumn = table.getColumn(4).getRangeBetweenHeaderAndTotal();
  combinedColumn.setFormula(`=[@[${nameOfColumn2}]]*[@[${nameOfColumn3}]]`);
}
```

### Before the script

| Month | Price | Units Sold | Total |
|--|--|--|--|
| Jan | 45 | 5 |  |
| Feb | 45 | 3 |  |
| Mar | 45 | 6 |  |

### After the script

| Month | Price | Units Sold | Total |
|--|--|--|--|
| Jan | 45 | 5 | 225 |
| Feb | 45 | 3 | 135 |
| Mar | 45 | 6 | 270 |
