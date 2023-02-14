---
title: Differences between Office Scripts and VBA macros
description: The behavior and API differences between Office Scripts and Excel VBA macros.
ms.date: 02/13/2023
ms.localizationpriority: medium
---

# Differences between Office Scripts and VBA macros

Office Scripts and VBA macros have a lot in common. They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings. Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.

The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed for secure, cross-platform, cloud-based solutions.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="A four-quadrant diagram showing the areas of focus for different Office extensibility solutions. Both Office Scripts and VBA macros are designed to help end users create solutions. Office Scripts are built for cross-platform experiences and collaboration, whereas VBA is for the desktop.":::

This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts. Since Office Scripts are only available for Excel, that is the only host being discussed here.

## Platform and ecosystem

The following table shows which features are supported by which platforms and products.

[!INCLUDE [Platform support table](../includes/platform-support-table.md)]

VBA is designed to be desktop-centric. VBA can interact with a user's desktop to connect with similar technologies, such as COM and OLE. However, VBA has no convenient way to call out to the internet. Office Scripts use a universal runtime for JavaScript. This gives consistent behavior and accessibility, regardless of the machine being used to run the script. They can also make calls to [a limited set of web services](../develop/external-calls.md).

## Security

VBA macros have the same security clearance as Excel. This gives them full access to your desktop. Office Scripts only have access to the workbook, not the machine hosting the workbook. Additionally, no JavaScript authentication tokens can be shared with scripts. This means the script has neither the tokens of the signed-in user nor are there any API capabilities for signing in to an external service, so they are unable to use existing tokens to make external calls on behalf of the user.

Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates. This lack of granularity makes it hard to isolate a single bad actor. Currently, Office Scripts can be off for an entire tenant, on for an entire tenant, or on for a group of users in a tenant. Admins also have control over who can share scripts with others and who can use scripts in Power Automate.

## Coverage

Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client. Office Scripts cover nearly all of the scenarios for Excel on the web. Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.

Office Scripts don't support Excel-level [events](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects). Scripts are only run when a user manually starts them or when a Power Automate flow calls the script.

## Power Automate

VBA doesn't have a Power Automate connector. All supported VBA scenarios involve a user attending to the macro's execution.

Office Scripts can be run through Power Automate. Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel. Try the [Call scripts from a manual Power Automate flow](../tutorials/excel-power-automate-manual.md) tutorial to start learning about Power Automate. You can also check out the [Automated task reminders](scenarios/task-reminders.md) sample to see Office Scripts connected to Teams through Power Automate in a real-world scenario.

## See also

- [Office Scripts in Excel](../overview/excel.md)
- [Run Office Scripts with Power Automate](../develop/power-automate-integration.md)
- [Differences between Office Scripts and Office Add-ins](add-ins-differences.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Excel VBA reference](/office/vba/api/overview/excel)
