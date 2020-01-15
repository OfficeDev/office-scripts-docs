---
title: 'Using built-in JavaScript libraries in Office Scripts'
description: 'How to call built-in JavaScript APIs from an Office Script in Excel on the web.'
ms.date: 01/14/2020
localization_priority: Normal
---

# Using built-in JavaScript libraries in Office Scripts

JavaScript has several built-in objects any JavaScript code can use. The [TypeScript](../overview/code-editor-environment.md) of Office Scripts is a superset of JavaScript and also includes these objects. This article focuses on a few select objects and how they integrate with an Excel workbook through a script. Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) contains a complete list of these objects.

## Date

The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script. `Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.

The following code shows how to add the current date to the worksheet. Note that because the string format matches that of an expected 