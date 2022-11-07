---
title: When to use Power Query or Office Scripts
description: The scenarios that are best suited for both the Power Query and Office Scripts platforms.
ms.date: 10/01/2022
ms.localizationpriority: medium
---

# When to use Power Query or Office Scripts

[Power Query](https://powerquery.microsoft.com) and Office Scripts are both powerful automation solutions for Excel. Both solutions let Excel users clean and transform data in workbooks. A single Power Query or Office Script can be refreshed and rerun on new data to produce consistent results, which saves you time and lets you work with the resulting information faster.

This article provides a general overview of when you might favor one platform over the other. In general, Power Query is good for pulling and transforming data from large, external data sources and Office Scripts are good for quick, Excel-centric solutions and [Power Automate integrations](../develop/power-automate-integration.md).

## Large data sources and data retrieval: Power Query

We recommend Power Query when dealing with data sources from supported platforms.

Power Query has [built-in data connections](https://powerquery.microsoft.com/connectors/) to hundreds of sources. Power Query is specially designed for data retrieval, transformation, and combination tasks. When you need data from one of those sources, Power Query gives you a no-code way of bringing that data into Excel in the shape you need.

These Power Query connections are designed for large datasets. They do not have the same [transfer limits](../testing/platform-limits.md) as Power Automate or Excel for the web.

Office Scripts offer a lightweight solution for smaller data sources or data sources not covered by Power Query connectors. This includes [using `fetch` or REST APIs](../develop/external-calls.md) or getting information from ad-hoc data sources, such as a [Teams adaptive card](../resources/scenarios/task-reminders.md).

## Formatting, visualizations, and programmatic control: Office Scripts

We recommend Office Scripts when your needs go beyond data importing and transformation.

Nearly everything you can do manually through the Excel UI is doable with Office Scripts. They're great for applying consistent formatting to workbooks. Scripts create charts, PivotTables, shapes, images, and other worksheet visualizations. Scripts also give you precise control over the positions, sizes, colors, and other attributes of these visualizations.

The inclusion of TypeScript code gives you a high degree of customization. Programmatic control logic like `if...else` statements makes your script robust. This lets you do things like conditionally read data without relying on complex Excel formulas, or scan the workbook for unexpected changes before changing the workbook.

Formatting can be applied with Power Query through Excel [templates](https://templates.office.com/power-query-tutorial-tm11414620). However, templates are updated at the individual or organization level, whereas Office Scripts offer more granular access control.

## Power Automate integrations

Office Scripts offer more options for Power Automate integration. Scripts are tailored to your solutions. You define the [input and output of the script](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), so it works with any other connector or data in the flow. The following screenshot shows an example Power Automate flow that passes data from a Teams Adaptive Card to an Office Script.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="A screenshot that shows the Excel Online (Business) connector in the flow designer. The connector is using the Run script action to take input from a Teams Adaptive Card and provide it to a script.":::

Power Query is used in the [SQL Server](https://powerquery.microsoft.com/flow/) Power Automate connector. The [Transform data using Power Query](/connectors/sql/#transform-data-using-power-query) action lets you build a query in Power Automate. While this is a powerful tool for use with SQL Server, it does limit Power Query to that input source, as shown in the following flow screenshot.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="A screenshot that shows the SQL Server connector in the flow designer. The connector is using the Transform data using Power Query action.":::

## Platform dependencies

Office Scripts is currently available for Excel for Windows and Excel on the web. The full Power Query experience is currently [only available for Excel for Windows](/power-query/power-query-what-is-power-query#where-can-you-use-power-query). Both can be used through Power Automate, which lets the flow work with Excel workbooks stored in OneDrive.

## See also

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query with Excel](https://powerquery.microsoft.com/excel/)
- [Run Office Scripts with Power Automate](../develop/power-automate-integration.md)
