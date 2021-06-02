---
title: 'Differences between Office Scripts and Office Add-ins'
description: 'The behavior and API differences between Office Scripts and Office Add-ins.'
ms.date: 06/02/2021
localization_priority: Normal
---

# Differences between Office Scripts and Office Add-ins

Office Add-ins and Office Scripts have a lot in common. They both offer automated control of an Excel workbook through JavaScript APIs. However, Office Scripts are designed to be quickly made by anyone looking to improve their workflow. Whereas Office Add-ins are larger-scale integrations for an Office host application.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="A four-quadrant diagram showing the focus areas for different Office extensibility solutions. Both Office Scripts and Office Web Add-ins are focused on the web and collaboration, but Office Scripts cater to end users (whereas Office Web Add-ins target professional developers)":::

Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open. This means the add-ins maintain state during a session, whereas Office Scripts don't maintain an internal state between runs. If you find that your Excel extension needs exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.

The rest of this article describes on the main differences between Office Add-ins and Office Scripts.

## Platform Support

Office Add-ins are cross-platform. They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each. Any exception to this is noted in the documentation of the individual API.

Office Scripts are currently only supported by for Excel on the web. All recording, editing, and running is done on the web platform.

## APIs

While the Office JavaScript APIs for Office Add-ins and the Office Scripts APIs share some functionality, they are different platforms. The Office Scripts APIs are an optimized, synchronous subset of the Excel JavaScript API model. The major difference is usage of the `load`/`sync` paradigm with add-ins. Additionally, add-ins offer APIs for events and a broader set of functionality outside of Excel, known as the Common APIs.

### Events

Office Scripts do not support workbook-level [events](/office/dev/add-ins/excel/excel-add-ins-events). Scripts are either triggered by users pressing the **Run** button for a script or through Power Automate. Every script runs the code in a single `main` method, then ends.

### Common APIs

Office Scripts cannot use [Common APIs](/javascript/api/office). If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.

## See also

- [Office Scripts in Excel on the web](../overview/excel.md)
- [Differences between Office Scripts and VBA macros](vba-differences.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Build an Excel task pane add-in](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
