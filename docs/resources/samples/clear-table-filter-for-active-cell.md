---
title: Remove table column filters
description: Learn how to clear table column filter based on active cell location.
ms.date: 07/15/2022
ms.localizationpriority: medium
---

# Remove table column filters

This sample remove the filters from a table column, based on the active cell location. The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.

If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.

## Sample Excel file

Download <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> for a ready-to-use workbook. Add the following script to try the sample yourself!

## Sample code: Clear table column filter based on active cell

The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
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

## Before clearing column filter (notice the active cell)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="An active cell before clearing column filter.":::

## After clearing column filter

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="An active cell after clearing column filter.":::
