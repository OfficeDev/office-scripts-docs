---
title: 'Clear table column filter based on active cell location'
description: 'Learn how to clear table column filter based on active cell location.'
ms.date: 05/03/2021
localization_priority: Normal
---

# Clear table column filter based on active cell location

This sample clears the table column filter based on the active cell location. The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.

If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.

_Before clearing column filter (notice the active cell)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="An active cell before clearing column filter":::

_After clearing column filter_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="An active cell after clearing column filter":::

## Sample code: Clear table column filter based on active cell

The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table. For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
