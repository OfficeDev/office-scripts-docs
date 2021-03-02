---
title: 'Copy Excel tables into new master table'
description: 'Learn how to use Office Scripts to combine data from multiple Excel tables into a single table.'
ms.date: 03/02/2021
localization_priority: Normal
---

# Copy Excel tables into new master table

Shows how to combine data that resides in Excel tables into a new master table with all rows included. It creates a new worksheet and combines all rows into a single table.

There are two variations of this script:

1. The first script combines all tables in the Excel file. It assumes that all tables in the Excel files have same structure.
1. The second script selectively gets tables within a set of worksheets. It assumes that all the tables in those select worksheets have the same structure.

## Video link

[![Watch step-by-step video on how to combine data from multiple Excel tables into a single table](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "Step-by-step video on how to combine data from multiple Excel tables into a single table")

## Office Scripts

This solution showcases two Office Scripts:

1. [Copy all tables in an Excel file to a new master table](#copy-all-tables-in-excel-to-a-new-master-table)
1. [Copy tables from select worksheets into a new master table](#copy-tables-from-select-worksheets-into-a-new-master-table)

### Copy all tables in Excel to a new master table

Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!

```ts
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

### Copy tables from select worksheets into a new master table

Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!

```ts
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```
