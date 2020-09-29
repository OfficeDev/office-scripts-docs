---
title: 'External API call support in Office Scripts'
description: 'Support and guidance for making external API calls in an Office Script.'
ms.date: 09/24/2020
localization_priority: Normal
---

# External API call support in Office Scripts

The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API). However, these calls can be run under the right circumstances. External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).

Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase. This is due how the JavaScript runtime manages interacting with the workbook. The script may end before the API call completes (or its `Promise` is fully resolved). As such, do not rely on external APIs for critical script scenarios.

> [!CAUTION]
> External calls may result in sensitive data being exposed to undesirable endpoints. Your admin can establish firewall protection against such calls.

## Definition files for external APIs

The definition files for external APIs aren't included with Office Scripts. The use of such APIs generates compile-time errors for missing definitions. The APIs still run (though only when run through the Excel client), as shown in the following script:

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## External calls from Power Automate

Any external API calls fail when a script is run with Power Automate. This is a behavioral difference between running a script through the Excel client and through Power Automate. Be sure to check your scripts for such references before building them into a flow.

> [!WARNING]
> The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies. However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls. For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts. Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).

## See also

- [Using built-in JavaScript objects in Office Scripts](javascript-objects.md)