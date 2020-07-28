---
title: 'Differences between Office Scripts and VBA macros'
description: 'The behavior and API differences between Office Scripts and Excel VBA macros.'
ms.date: 06/30/2020
localization_priority: Normal
---

# Differences between Office Scripts and VBA macros

Office Scripts and VBA macros have a lot in common. They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings. Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.
The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed with cross-platform support and security as the guiding principles. Currently, Office Scripts are only supported in Excel on the web.

![A four-quadrant diagram showing the areas of focus for different Office extensibility solutions. Both Office Scripts and VBA macros are designed to help end users create solutions, but Office Scripts are built for the web and collaboration (whereas VBA is for the desktop).)](../images/office-programmability-diagram.png)

This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts. Since Office Scripts are only available for Excel, that is the only host being discussed here.

## Platform and ecosystem

VBA is designed for the desktop and Office Scripts are designed for the web. VBA can interact with a user's desktop to connect with similar technologies, such as COM and OLE. However, VBA has no convenient way to call out to the internet.

Office Scripts use a universal runtime for JavaScript. This gives consistent behavior and accessibility, regardless of the machine being used to run the script. They can also make calls to other web services.

## Security

VBA macros have the same security clearance as Excel. This gives them full access to your desktop. Office Scripts only have access to the workbook, not the machine hosting the workbook. Additionally, no JavaScript authentication tokens can be shared with scripts, so scripts can never authenticate with an external service.

Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates. This lack of granularity makes it hard to isolate a single bad actor. Currently, Office Scripts are either on or off for a tenant. However, we are working to give admins more control over individual scripts and script creators.

## Coverage

Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client. Office Scripts cover nearly all of the scenarios for Excel on the web. Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.

## Power Automate

Office Scripts can be run through Power Automate. Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel. This means that as long as your workbook is stored in OneDrive (and accessible to Power Automate), a flow can run your scripts regardless of whether you and your organization use Excel's desktop, Mac, or web client.

VBA doesn't have a Power Automate connector. All supported VBA scenarios involved a user attending to the macro's execution.

## See also

- [Office Scripts in Excel on the web](../overview/excel.md)
- [Differences between Office Scripts and Office Add-ins](add-ins-differences.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Excel VBA reference](/office/vba/api/overview/excel)
