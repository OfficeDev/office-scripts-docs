---
title: 'Differences between Office Scripts and Office Add-ins'
description: 'The behavior and API differences between Office Scripts and Office Add-ins.'
ms.date: 11/07/2019
localization_priority: Normal
---

# Differences between Office Scripts and Office Add-ins

Office Add-ins and Office Scripts have a lot in common. They both offer automated control of an Excel workbook through the Excel namespace of the Office JavaScript API. However, Office Scripts are more limited in their scope.

Scripts run to completion with a manual button press, whereas Add-ins rely on user-interaction and persist while the workbook is being used. If you find your Excel extension needs exceed the Scripting platform’s capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more.

This rest of this article focuses on the programmatic differences between Office Add-ins and Excel Scripts.

## Platform Support

Add-ins are cross-platform. They are intended to work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each. Any exception to this is noted in the documentation of the individual API.

Excel Scripts are only available for Excel on the web. All recording, editing, and running is done on the web platform.

## APIs

Office Scripts support most of the Excel APIs. This means there is a lot of functionality overlap between the two platforms. There are two exceptions:

- [Events](/office/dev/add-ins/excel/excel-add-ins-events)
- [Common APIs](/javascript/api/office)

Office Scripts do not support events. Every script runs the code in a single `main` method, then ends. It does not reactivate when events are triggered, and thus, cannot register events.
All of the Common APIs are excluded from Office Scripts. If you need authentication, dialog windows, or other such features, you likely require an Office Add-in instead of a script.

## See also

- [Overview: Office Scripts in Excel on the web](../overview/overview.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)