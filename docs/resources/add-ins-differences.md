---
title: Differences between Office Scripts and Office Add-ins
description: The behavior and API differences between Office Scripts and Office Add-ins.
ms.date: 02/04/2022
ms.localizationpriority: medium
---

# Differences between Office Scripts and Office Add-ins

Understand the differences between Office Scripts and Office Add-ins to know when to use each one. Office Scripts are designed to be quickly made by anyone looking to improve their workflow. Office Add-ins integrate with the Office UI for a more interactive experience through ribbon buttons and task panes. Office Add-ins can also expand built-in Excel functions by providing custom functions.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="A four-quadrant diagram showing the focus areas for different Office extensibility solutions. Both Office Scripts and Office Web Add-ins are focused on the web and collaboration, but Office Scripts cater to end users (whereas Office Web Add-ins target professional developers).":::

Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins continue running depending on how they are configured. For example, you can configure an Office Add-in to continue running even when its task pane is closed. This means that Office Add-ins maintain state during a session, whereas Office Scripts don't maintain an internal state between runs. If the solution you are building requires a maintained state, you should visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.

The rest of this article describes on the main differences between Office Add-ins and Office Scripts.

## Platform Support

Office Add-ins are cross-platform. They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each. Any exception to this is noted in the documentation of the individual API.

Office Scripts are currently only supported by for Excel on the web. All recording, editing, and script management is done on the web platform.

## APIs

While the Office JavaScript APIs for Office Add-ins and the Office Scripts APIs share some functionality, they are different platforms. The Office Scripts APIs are an optimized, synchronous subset of the Excel JavaScript API model. The major difference is usage of the `load`/`sync` paradigm with add-ins. Additionally, add-ins offer APIs for events and a broader set of functionality outside of Excel, known as the Common APIs.

### Events

Office Scripts do not support workbook-level [events](/office/dev/add-ins/excel/excel-add-ins-events). Scripts are either triggered by users selecting the **Run** button for a script or through Power Automate. Every script runs the code in a single `main` function, then ends.

### Common APIs

Office Scripts cannot use [Common APIs](/javascript/api/office). If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.

## See also

- [Office Scripts in Excel](../overview/excel.md)
- [Differences between Office Scripts and VBA macros](vba-differences.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Build an Excel task pane add-in](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
