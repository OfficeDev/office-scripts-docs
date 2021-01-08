---
title: 'External API call support in Office Scripts'
description: 'Support and guidance for making external API calls in an Office Script.'
ms.date: 01/05/2021
localization_priority: Normal
---

# External API call support in Office Scripts

Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase. As such, do not rely on external APIs for critical script scenarios.

Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).

> [!CAUTION]
> External calls may result in sensitive data being exposed to undesirable endpoints. Your admin can establish firewall protection against such calls.

## Working with `fetch`

The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services. It is an `async` API, so you will need to adjust the `main` signature of your script. Make the `main` function `async` and have it return a `Promise<void>`. You should also be sure to `await` the `fetch` call and `json` retrieval. This ensures those operations complete before the script ends.

The following script uses `fetch` to retrieve JSON data from the test server in the given URL.

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.

## External calls from Power Automate

Any external API calls fail when a script is run with Power Automate. This is a behavioral difference between running a script through the Excel client and through Power Automate. Be sure to check your scripts for such references before building them into a flow.

> [!WARNING]
> External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies. However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls. For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts. Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).

## See also

- [Using built-in JavaScript objects in Office Scripts](javascript-objects.md)
- [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md)
