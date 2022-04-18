---
title: 'Work with PivotTables in Office Scripts'
description: 'The object model for PivotTables in the Office Scripts JavaScript API'
ms.date: 04/15/2022
ms.localizationpriority: medium
---

# Work with PivotTables in Office Scripts

PivotTables let you analyze large collections of data faster. With their power comes complexity. The Office Scripts APIs let you customize a PivotTable to suit your needs, but the scope of the API set makes getting started a challenge. This article demonstrates how to perform common PivotTable tasks and explains the important classes and methods that are frequently used.

> [!NOTE]
> To better understand context for the terms used by the APIs, please read Excel's PivotTable documentation first, starting with [Create a PivotTable to analyze worksheet data](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## Object model

:::image type="content" source="../images/pivottable-object-model.png" alt-text="A simplified picture of the classes, methods, and properties used when working with PivotTables.":::

The [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) is the central object for PivotTables in the Office JavaScript API.

- The [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) object has a collection of all the [PivotTables](/javascript/api/office-scripts/excelscript/excelscript.pivottable). Each [Worksheet also contains a PivotTable collection that's local to that sheet.
- A [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) contains [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). A hierarchy can be thought of as a column in a table.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) can be added as rows or columns ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), data ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)), or filters ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)).
- Each [PivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) contains one and only one [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). PivotTable structures outside of Excel may contain multiple fields pet hierarchy, so this design exists to support future options. For Office Scripts, fields and hierarchies map to the same information.
- A [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) contains multiple [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Each PivotItem is a unique value in the field. Think of each item as a value in the table column. Items could also be aggregated values, such as sums, if the field is being used for data.
- The [PivotLayout defines how the [PivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) and [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem) are displayed.
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) filter data from the [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) using different criteria.

Let's look at how these relationships work in practice. The following data describes fruit sales from various farms. It will be the example throughout this article. Use <a href="pivottable-sample.xlsx">pivottable-sample.xlsx</a> to follow along with the samples in the article.

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="A collection of fruit sales of different types from different farms.":::

## Create a PivotTable with fields

PivotTables are created with references to existing data. Both ranges and tables can be the source for a PivotTable. They also need a place to exist in the workbook. Since the size of PivotTable is dynamic, only the upper-left corner of the destination range is specified.

The following code snippet creates a PivotTable based on a range of data. The PivotTable has no hierarchies, so the data is not yet grouped in any way.

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="A PivotTable named 'Farm Pivot' with no hierarchies.":::

### Hierarchies and fields

PivotTables are organized through hierarchies. There are four types of hierarchies.

- **Row**: Displays items in horizontal rows.
- **Column**: Displays items in vertical columns.
- **Data**: Displays aggregates of values based on the rows and columns.
- **Filter**: Add or removes items from the PivotTable.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="A PivotTable showing the total sales of different fruit based on the farm they came from.":::

## Layout ranges

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="A diagram showing which sections of a PivotTable are returned by the layout's get range functions.":::

## Filters and slicers

:::image type="content" source="../images/slicer.png" alt-text="A slicer filtering data on a PivotTable.":::

## Change aggregation function and calculations

:::image type="content" source="../images/pivottable-showas-percentage.png" alt-text="A PivotTable showing the percentages of fruit sales relative to the grand total for both individual farms and individual fruit types within each farm.":::

:::image type="content" source="../images/pivottable-showas-differencefrom.png" alt-text="A PivotTable showing the differences of fruit sales between 'A Farms' and the others. This shows both the difference in total fruit sales of the farms and the sales of types of fruit. If 'A Farms' did not sell a particular type of fruit, '#N/A' is displayed.":::
