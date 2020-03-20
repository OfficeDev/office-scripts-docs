---
title: 'Differences between Office Scripts and Office Add-ins'
description: 'The behavior and API differences between Office Scripts and Office Add-ins.'
ms.date: 12/12/2019
localization_priority: Normal
---

# Differences between Office Scripts and Office Add-ins

Office Add-ins and Office Scripts have a lot in common. They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API. However, Office Scripts are more limited in their scope.

![A four-quadrant diagram showing the areas of focus for different Office extensibility solutions. Both Office Scripts and Office Web Add-ins are focused on web and collaboration, but Office Scripts are more end-user approachable (whereas Office Web Add-ins are targeted at a professional developer audience).)](../images/office-prgrammability-diagram.png)

Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins rely on user interaction and persist while the workbook is in use. If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.

The rest of this article describes on the main differences between Office Add-ins and Office Scripts.

## Platform Support

Office Add-ins are cross-platform. They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each. Any exception to this is noted in the documentation of the individual API.

Office Scripts are currently only supported by for Excel on the web. All recording, editing, and running is done on the web platform.

## APIs

Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms. There are two exceptions: events and Common APIs.

### Events

Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events). Every script runs the code in a single `main` method, then ends. It does not reactivate when events are triggered, and thus, cannot register events.

### Common APIs

Office Scripts cannot use [Common APIs](/javascript/api/office). If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.

## See also

- [Office Scripts in Excel on the web](../overview/excel.md)
- [Differences between Office Scripts and VBA macros](vba-differences.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Build an Excel task pane add-in](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)