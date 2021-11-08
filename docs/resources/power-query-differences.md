---
title: 'When to use Power Query or Office Scripts'
description: 'The scenarios that are best suited for both the Power Query and Office Scripts platforms.'
ms.date: 11/05/2021
ms.localizationpriority: medium
---

# When to use Power Query or Office Scripts

[Power Query](https://powerquery.microsoft.com) and Office Scripts are both powerful automation solutions for Excel. This article provides a general overview of when you might favor one platform over the other. In general, Power Query is good for pulling and transforming data from large, external data sources and Office Scripts are good for quick, Excel-centric solutions and Power Automate integrations.

## Large data sources: Power Query

We recommend Power Query when dealing with large data sources in supported platforms.

Power Query has [built-in data connections](https://powerquery.microsoft.com/connectors/) to hundreds of sources. These connections are designed for large datasets. They do not have the same transfer limits as Power Automate or Excel for the web.

For smaller data sources or data sources not covered by Power Query connectors, Office Scripts offer a lightweight solution. This includes [using `fetch` or REST APIs](../develop/external-calls.md) or getting information from ad-hoc data sources, such as a [Teams adaptive card](../resources/scenarios/task-reminders.md).

## Formatting and programmatic control: Office Scripts

We recommend Office Scripts for formatting workbooks, particularly in scenarios with conditional logic.

Office Scripts are great for applying a standard format to workbooks. The inclusion of TypeScript code gives you a high degree of customization.

Formatting can be applied with Power Query through [templates](https://templates.office.com/power-query-tutorial-tm11414620). However, templates cannot be automatically updated for your entire organization. New versions of templates must be distributed and adopted, whereas Office Scripts are attached ot workbooks and automatically updated when the script creator edits them.

## Team deployment: Office Scripts

We recommend using Office Scripts if you need to deploy and maintain automation solutions across your team or organization.

When [attached to workbooks](../overview/excel.md#share-scripts), Office Scripts give all users of the workbook shared access to the script. All users are running the same script and cannot be on different versions. Updates are automatically made after the creator edits the script.

## Power Automate integrations

Office Scripts offer more options for Power Automate integration. Scripts are tailored to your solutions. You define the [input and output of the script](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), so it works with any other connector or data in the flow.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="The Excel Online (Business) connector in the flow designer showing the Office Scripts option.":::

Power Query is used in the [SQL Server](https://powerquery.microsoft.com/flow/) Power Automate connector. The [Transform data using Power Query](/connectors/sql/#transform-data-using-power-query) action lets you build a query in Power Automate. While this is a powerful tool for use with SQL Server, it does limit Power Query to that input source.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="The SQL Server connector in the flow designer showing the Power Query option.":::

## Platform dependencies

Office Scripts is currently only for Excel on the web. Power Query is currently only for Excel on Desktop. Both can be used through Power Automate, which lets the flow work with Excel workbooks stored in OneDrive.

## See also

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query with Excel](https://powerquery.microsoft.com/excel/)
- [Run Office Scripts with Power Automate](../develop/power-automate-integration.md)
