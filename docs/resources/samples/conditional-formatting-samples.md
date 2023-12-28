---
title: Conditional formatting samples
description: A collection of Office Scripts that use different Excel conditional formatting options.
ms.date: 12/28/2023
ms.localizationpriority: medium
---

# Conditional formatting samples

Conditional formatting in Excel applies formatting to cells based on specific conditions or rules. This feature helps you visually highlight important data, identify trends, and analyze patterns in your spreadsheet. With conditional formatting, you'll quickly use color scales, data bars, and icon sets, to dynamically format your data. This page contains a collection of Office Scripts that demonstrate various conditional formatting options.

This sample workbook contains worksheets ready to test with the sample scripts.

> [!div class="nextstepaction"]
> [Download the sample workbook](conditional-formatting-samples.xlsx)

## Cell value

[Cell value conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.cellvalueconditionalformat) applies a format to every cell that contains a value meeting a given criteria. This helps quickly spot important data points.

The following sample applies cell value conditional formatting to a range. Any value less than 60 will have the cell's fill color changed and the font made italic.

:::image type="content" source="../../images/conditional-formatting-sample-cell-value.png" alt-text="A list of scores with every cell that contains a value under 60 formatted to have a yellow fill and italic text.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("CellValue");
    const ratingColumn = sheet.getRange("B2:B12");
    sheet.activate();
    
    // Add cell value conditional formatting.
    const cellValueConditionalFormatting =
        ratingColumn.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue).getCellValue();
    
    // Create the condition, in this case when the cell value is less than 60
    let rule: ExcelScript.ConditionalCellValueRule = {
        formula1: "60",
        operator: ExcelScript.ConditionalCellValueOperator.lessThan
    };
    cellValueConditionalFormatting.setRule(rule);
    
    // Set the format to apply when the condition is met.
    let format = cellValueConditionalFormatting.getFormat();
    format.getFill().setColor("yellow");
    format.getFont().setItalic(true);
}
```

## Color scale

[Color scale conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.colorscaleconditionalformat) applies a color gradient across a range. The cells with the minimum and maximum values of the range use the colors specified, with other cells scaled proportionally. An optional midpoint color provides more contrast.

This following sample applies a red, white, and blue color scale to the selected range.

:::image type="content" source="../../images/conditional-formatting-sample-color-scale.png" alt-text="A table of temperatures with the lower values colored blue and the higher ones colored red.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("ColorScale");
    const dataRange = sheet.getRange("B2:M13");
    sheet.activate();

    // Create a new conditional formatting object by adding one to the range.
    const conditionalFormatting = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);

    // Set the colors for the three parts of the scale: minimum, midpoint, and maximum.
    conditionalFormatting.getColorScale().setCriteria({
        minimum: {
            color: "#5A8AC6", /* A pale blue. */
            type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue
        },
        midpoint: {
            color: "#FCFCFF", /* Slightly off-white. */
            formula: '=50', type: ExcelScript.ConditionalFormatColorCriterionType.percentile
        },
        maximum: {
            color: "#F8696B", /* A pale red. */
            type: ExcelScript.ConditionalFormatColorCriterionType.highestValue
        }
    });
}
```

## Data bar

[Data bar conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.databarconditionalformat) adds a partially-filled bar in the background of a cell. The fullness of the bar is defined by the value in the cell and the range specified by the format.

The following sample creates data bar conditional formatting on the selected ranged. The scale of the data bar goes from 0 to 1200.

:::image type="content" source="../../images/conditional-formatting-sample-data-bar.png" alt-text="A table of values with data bars showing their value compared to 1200.":::

```typescript

function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("DataBar");
    const dataRange = sheet.getRange("B2:D5");
    sheet.activate();

    // Create new conditional formatting on the range.
    const format = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.dataBar);
    const dataBarFormat = format.getDataBar();

    // Set the lower bound of the data bar formatting to be 0.
    const lowerBound: ExcelScript.ConditionalDataBarRule = {
        type: ExcelScript.ConditionalFormatRuleType.number,
        formula: "0"
    };
    dataBarFormat.setLowerBoundRule(lowerBound);

    // Set the upper bound of the data bar formatting to be 1200.
    const upperBound: ExcelScript.ConditionalDataBarRule = {
        type: ExcelScript.ConditionalFormatRuleType.number,
        formula: "1200"
    };
    dataBarFormat.setUpperBoundRule(upperBound);
}
```

## Icon set

[Icon set conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.iconsetconditionalformat) adds icons to each cell in a range. The icons come from a specified set. Icons are applied based on an ordered array of criteria, with each criterion mapping to a single icon.

The following sample applies the "three traffic light" icon set conditional formatting to a range.

:::image type="content" source="../../images/conditional-formatting-sample-icon-set.png" alt-text="A table of scores with red lights next to low values, yellow lights next to medium values, and green lights next to high values.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("IconSet");
    const dataRange = sheet.getRange("B2:B12");
    sheet.activate();

    // Create icon set conditional formatting on the range.
    const conditionalFormatting = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.iconSet);

    // Use the "3 Traffic Lights (Unrimmed)" set.
    conditionalFormatting.getIconSet().setStyle(ExcelScript.IconSet.threeTrafficLights1);
    conditionalFormatting.getIconSet().setCriteria([
      { // Use the red light as the default for positive values.
        formula: '=0', operator: ExcelScript.ConditionalIconCriterionOperator.greaterThanOrEqual,
        type: ExcelScript.ConditionalFormatIconRuleType.number
      },
      { // The yellow light is applied to all values 6 and greater. The replaces the red light when applicable.
        formula: '=6', operator: ExcelScript.ConditionalIconCriterionOperator.greaterThanOrEqual,
        type: ExcelScript.ConditionalFormatIconRuleType.number
      },
      { // The green light is applied to all values 8 and greater. As with the yellow light, the icon is replaced when the new criteria is met.
        formula: '=8', operator: ExcelScript.ConditionalIconCriterionOperator.greaterThanOrEqual,
        type: ExcelScript.ConditionalFormatIconRuleType.number
      }
    ]);
}
```

## Preset

[Preset conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.presetcriteriaconditionalformat) applies a specified format to a range based on common scenarios, such as blank cells and duplicate values. The full list of preset criteria is provided by the [ConditionalFormatPresetCriterion](/javascript/api/office-scripts/excelscript/excelscript.conditionalformatpresetcriterion) enum.

The following sample gives a yellow fill to any blank cell in the range.

:::image type="content" source="../../images/conditional-formatting-sample-preset.png" alt-text="A table with blank values highlighted with yellow fills.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("Preset");
    const dataRange = sheet.getRange("B2:D5");
    sheet.activate();
    
    // Add new conditional formatting to that range.
    const conditionalFormat = dataRange.addConditionalFormat(
    ExcelScript.ConditionalFormatType.presetCriteria);
    
    // Set the conditional formatting to apply a yellow fill.
    const presetFormat = conditionalFormat.getPreset();
    presetFormat.getFormat().getFill().setColor("yellow");
    
    // Set a rule to apply the conditional format when cells are left blank.
    const blankRule: ExcelScript.ConditionalPresetCriteriaRule = {
        criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks
    };
    presetFormat.setRule(blankRule);
}
```

## Text comparison

[Text comparison conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.textconditionalformat) formats cells based on their text content. The formatting is applied when the text begins with, contains, ends with, or doesn't contain the given substring.

The following sample marks any cell in the range that contains the text "review".

:::image type="content" source="../../images/conditional-formatting-sample-text-comparison.png" alt-text="A table with status entries where any cell containing the word 'review' has a red fill.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("TextComparison");
    const dataRange = sheet.getRange("B2:B6");
    sheet.activate();

    // Add conditional formatting based on the text in the cells.
    const textConditionFormat = dataRange.addConditionalFormat(
        ExcelScript.ConditionalFormatType.containsText).getTextComparison();

    // Set the conditional format to provide a light red fill.
    textConditionFormat.getFormat().getFill().setColor("#F8696B");

    // Apply the condition rule that the text contains with "review".
    const textRule: ExcelScript.ConditionalTextComparisonRule = {
        operator: ExcelScript.ConditionalTextOperator.contains,
        text: "review"
    };
    textConditionFormat.setRule(textRule);
}
```

## Top/bottom

[Top/bottom conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.topbottomconditionalformat) marks the highest or lowest values in a range. The highs and lows are based on either raw values or percentages.

The following sample applies conditional formatting to show the two highest numbers in the range.

:::image type="content" source="../../images/conditional-formatting-sample-top-bottom.png" alt-text="A sales table that has the top two values highlighted with a green fill.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the range to format.
  const sheet = workbook.getWorksheet("TopBottom");
  const dataRange = sheet.getRange("B2:D5");
  sheet.activate();

    // Set the fill color to green for the top 2 values in the range.
    const topBottomFormat = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom).getTopBottom();
    topBottomFormat.getFormat().getFill().setColor("green");
    topBottomFormat.setRule({
      rank: 2, /* The numeric threshold. */
      type: ExcelScript.ConditionalTopBottomCriterionType.topItems /* The type of the top/bottom condition. */
    });
}
```

## Custom conditions

[Custom conditional formatting](/javascript/api/office-scripts/excelscript/excelscript.customconditionalformat) allows for complex formulas to define when formatting is applied. Use this when the other options aren't enough.

The following sample sets a custom conditional formatting on the selected range. A light-green fill is applied to a cell if the value is larger than the value in the row's previous column.

:::image type="content" source="../../images/conditional-formatting-sample-custom.png" alt-text="A row of a sales table. Values that are higher than the one to the left have a green fill.":::

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the range to format.
    const sheet = workbook.getWorksheet("Custom");
    const dataRange = sheet.getRange("B2:H2");
    sheet.activate();
    
    // Apply a rule for positive change from the previous column.
    let positiveChange = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    positiveChange.getCustom().getFormat().getFill().setColor("lightgreen");
    positiveChange.getCustom().getRule().setFormula(
        `=${dataRange.getCell(0, 0).getAddress()}>${dataRange.getOffsetRange(0, -1).getCell(0, 0).getAddress()}`
    );
}
```
