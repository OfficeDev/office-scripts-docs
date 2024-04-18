---
title: Work with PivotTables in Office Scripts
description: Learn about the object model for PivotTables in the Office Scripts JavaScript API.
ms.date: 04/18/2024
ms.localizationpriority: medium
---

# Work with PivotTables in Office Scripts

PivotTables let you analyze large collections of data quickly. With their power comes complexity. The Office Scripts APIs let you customize a PivotTable to suit your needs, but the scope of the API set makes getting started a challenge. This article demonstrates how to perform common PivotTable tasks and explains important classes and methods.

> [!NOTE]
> To better understand context for the terms used by the APIs, read Excel's PivotTable documentation first. Start with [Create a PivotTable to analyze worksheet data](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## Object model

:::image type="content" source="../images/pivottable-object-model.png" alt-text="A simplified picture of the classes, methods, and properties used when working with PivotTables.":::

The [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) is the central object for PivotTables in the Office Scripts API.

- The [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) object has a collection of all the [PivotTables](/javascript/api/office-scripts/excelscript/excelscript.pivottable). Each [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) also contains a PivotTable collection that's local to that sheet.
- A [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) contains [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). A hierarchy can be thought of as a column in a table.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) can be added as rows or columns ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), data ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)), or filters ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)).
- Each [PivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) contains exactly one [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). PivotTable structures outside of Excel may contain multiple fields per hierarchy, so this design exists to support future options. For Office Scripts, fields and hierarchies map to the same information.
- A [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) contains multiple [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Each PivotItem is a unique value in the field. Think of each item as a value in the table column. Items could also be aggregated values, such as sums, if the field is being used for data.
- The [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) defines how the [PivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) and [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem) are displayed.
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) filter data from the [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) using different criteria.

To look at how these relationships work in practice, start by downloading the sample workbook. That data describes fruit sales from various farms. It's the base for all the examples in this article. Run the sample scripts throughout the article to create and explore PivotTables.

> [!div class="nextstepaction"]
> [Download the sample workbook](pivottable-sample.xlsx)

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="A collection of fruit sales of different types from different farms.":::

## Create a PivotTable with fields

PivotTables are created with references to existing data. Both ranges and tables can be the source for a PivotTable. They also need a place to exist in the workbook. Since the size of a PivotTable is dynamic, only the upper-left corner of the destination range is specified.

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

PivotTables are organized through hierarchies. Those hierarchies are used to pivot data when added as a specific type of hierarchy. There are four types of hierarchies.

- **Row**: Displays items in horizontal rows.
- **Column**: Displays items in vertical columns.
- **Data**: Displays aggregates of values based on the rows and columns.
- **Filter**: Add or removes items from the PivotTable.

A PivotTable can have as many or as few of its fields assigned to these specific hierarchies. A PivotTable needs at least one data hierarchy to show summarized numerical data and at least one row or column to pivot that summary on. The following code snippet adds two row hierarchies and two data hierarchies.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="A PivotTable showing the total sales of different fruit based on the farm they came from.":::

## Layout ranges

Each part of the PivotTable maps to a range. This lets your script get data from the PivotTable for use later in the script or to be returned in a [Power Automate flow](power-automate-integration.md). These ranges are accessed through the [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) object acquired from `PivotTable.getLayout()`. The following diagram shows the ranges that are returned by the methods in `PivotLayout`.

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="A diagram showing which sections of a PivotTable are returned by the layout's get range functions.":::

### PivotTable total output

The location of the total row is based on the layout. Use [`PivotLayout.getBodyAndTotalRange`](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout#excelscript-excelscript-pivotlayout-getbodyandtotalrange-member(1)) and get the last row of the column to use the data from the PivotTable in your script.

The following sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).

:::image type="content" source="../images/sample-pivottable-grand-total-row.png" alt-text="A PivotTable showing fruit sales with the Grand Total row highlighted green.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  const pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  const pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  const pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  const grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## Filters and slicers

There are three ways to filter a PivotTable.

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### FilterPivotHierarchies

`FilterPivotHierarchies` add an additional hierarchy to filter every data row. Any row with an item that is filtered out is excluded from the PivotTable and its summaries. Since these filters are based on items, they only work on discrete values. If "Classification" is a filter hierarchy in the sample, users can select the values of "Organic" and "Conventional" for the filter. Similarly, if "Crates Sold Wholesale" is selected, the filter options would be the individual numbers, such as 120 and 150, instead of numerical ranges.

`FilterPivotHierarchies` are created with all values selected. This means that nothing is filtered until the user manually interacts with the filter control or a `PivotManualFilter` is set on the field belonging to the `FilterPivotHierarchy`.

The following code snippet adds "Classification" as a filter hierarchy.

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="A filter control that uses 'Classification' for a PivotTable.":::

### PivotFilters

The `PivotFilters` object is a collection of filters applied to a single field. Since each hierarchy has exactly one field, you should always use the first field in `PivotHierarchy.getFields()` when applying filters. There are four filter types.

- **Date filter**: Calendar date-based filtering.
- **Label filter**: Text comparison filtering.
- **Manual filter**: Custom input filtering.
- **Value filter**: Number comparison filtering. This compares items in the associated hierarchy with values in a specified data hierarchy.

Typically, only one of the four types of filters is created and applied to the field. If the script tries to use incompatible filters, an error is thrown with the text "The argument is invalid or missing or has an incorrect format."

The following code snippet adds two filters. The first is a manual filter that selects items in an existing "Classification" filter hierarchy. The second filter removes any farms that have fewer than 300 "Crates Sold Wholesale". Note that this filters out the "Sum" of those farms, not the individual rows from the original data.

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="A PivotTable after the value filter and manual filter were applied.":::

### Slicers

[Slicers](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) filter data in a PivotTable (or standard table). They are moveable objects in the worksheet that allow for quick filtering selections. A slicer operates in a similar fashion to the manual filter and `PivotFilterHierarchy`. Items from the `PivotField` are toggled to include or exclude them from the PivotTable.

The following code snippet adds a slicer for the "Type" field. It sets the selected items to be "Lemon" and "Lime", then moves the slicer 400 pixels to the left.

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="A slicer filtering data on a PivotTable.":::

### Value field settings for summaries

Change how the PivotTable summarizes and displays data with these settings. The field in each data hierarchy can display the data in different ways, such as percentages, standard deviations, and relative comparisons.

#### Summarize by

The default summarization of a data hierarchy field is as a sum. `DataPivotHierarchy.setSummarizeBy` lets you combine the data for each row or column in a different way. [`AggregationFunction`](/javascript/api/office-scripts/excelscript/excelscript.aggregationfunction) lists the all the available options.

The following code snippet changes "Crates Sold Wholesale" to show each item's standard deviation, instead of the sum.

```typescript
  const wholesaleSales = farmPivot.getDataHierarchy("Sum of Crates Sold Wholesale");
  wholesaleSales.setSummarizeBy(ExcelScript.AggregationFunction.standardDeviation);
```

#### Show values as

`DataPivotHierarchy.setShowAs` applies a calculation to the values of a data hierarchy. Instead of the default sum, you can show values or percentages relative to other parts of the PivotTable.

The following code snippet changes the display for "Crates Sold at Farm". The values will be shown as a percentage of the grand total for the field.

```typescript
  const farmSales = farmPivot.getDataHierarchy("Sum of Crates Sold at Farm");
  farmSales.setShowAs({
    calculation: ExcelScript.ShowAsCalculation.percentOfGrandTotal
  });
```

Some `ShowAsRule`s need another field or item in that field as a comparison. The following code snippet again changes the display for "Crates Sold at Farm". This time, the field will show each value's difference from the value of the "Lemons" in that farm row. If a farm has not sold any lemons, the field shows "#N/A".

```typescript
  const farmSales = farmPivot.getDataHierarchy("Sum of Crates Sold at Farm");

  const typeField = farmPivot.getRowHierarchy("Type").getFields()[0];
  farmSales.setShowAs({
    calculation: ExcelScript.ShowAsCalculation.differenceFrom,
    baseField: typeField, /* The field to use for the difference. */
    baseItem: typeField.getPivotItem("Lemon") /* The item within that field that is the basis of comparison for the difference. */
  });
  farmSales.setName("Difference from Lemons of Crates Sold at Farm");
```

## See also

- [Fundamentals for Office Scripts in Excel](scripting-fundamentals.md)
- [Office Scripts API reference](/javascript/api/office-scripts/overview)
